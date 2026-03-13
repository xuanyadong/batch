package com.tjexcel.batch.service;

import com.tjexcel.batch.config.BillOfLadingConfig;
import com.tjexcel.batch.util.OrderSplitUtil;
import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 提单批量生成服务
 * 从数据表读取每一行，替换占位符，填充内置表格（数量拆分后逐行）
 * <p>
 * 卡号全局对应：同一公司作为需方收货的卡号总量 = 作为供方发货的卡号总量，
 * 保证收货/发货的卡号能够一一对应。
 */
@Service
public class BillOfLadingService {

    private static final Logger log = LoggerFactory.getLogger(BillOfLadingService.class);

    private static final int PIECES_MULTIPLIER = 40;
    private static final int ALLOCATION_RETRY_LIMIT = 20;
    /**
     * 续收停止阈值：当卡累计量接近 split-max 时不再续收。
     * 例如 split-max=495 时，达到 450+ 即视为接近满额，不再补齐到 495。
     */
    private static final int STOP_FILL_GAP = 45;

    /** 表格列：产品名 规格型号 卡号 数量（吨） 件数 备注 */
    private static final String[] TABLE_HEADERS = {"产品名", "规格型号", "卡号", "数量（吨）", "件数", "备注"};

    private final BillOfLadingConfig config;

    public BillOfLadingService(BillOfLadingConfig config) {
        this.config = config;
    }

    public int generate() throws IOException {
        Path dataFile = Paths.get(config.getDataPath()).toAbsolutePath();
        Path templateFile = Paths.get(config.getTemplatePath()).toAbsolutePath();
        Path outputRoot = Paths.get(config.getOutputDir()).toAbsolutePath();

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("提单数据表不存在: " + dataFile);
        }
        if (!Files.exists(templateFile)) {
            throw new FileNotFoundException("提单模板不存在: " + templateFile);
        }

        Files.createDirectories(outputRoot);

        int totalSuccess = 0;

        try (InputStream is = Files.newInputStream(dataFile);
             Workbook workbook = openWorkbook(is, dataFile)) {
            int sheetCount = workbook.getNumberOfSheets();
            if (sheetCount <= 0) {
                log.warn("提单数据表没有可用 sheet，无提单可生成");
                return 0;
            }

            for (int s = 0; s < sheetCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                if (sheet == null) continue;
                String sheetName = sheet.getSheetName();
                String safeSheetName = (sheetName == null || sheetName.trim().isEmpty())
                        ? "sheet" + (s + 1)
                        : sheetName.trim().replaceAll("[\\\\/:*?\"<>|]", "_");
                Path sheetOutputDir = outputRoot.resolve(safeSheetName);
                Files.createDirectories(sheetOutputDir);

                List<Map<String, String>> rows = readDataSheet(sheet);
                log.info("从 {} 的 sheet[{}] 读取到 {} 行提单数据", dataFile, sheetName, rows.size());
                if (rows.isEmpty()) {
                    log.info("sheet[{}] 无有效提单数据，跳过生成", sheetName);
                    continue;
                }

                // 预计算：按 (公司, 规格型号) 建立卡号库存，并预分配每行的卡号明细（仅针对当前 sheet）
                CardAllocationContext ctx = buildCardAllocationContext(rows);

                Map<String, Integer> seqMap = new HashMap<>();
                int successCount = 0;
                Set<String> usedFileNames = new HashSet<>();
                for (int i = 0; i < rows.size(); i++) {
                    Map<String, String> rowData = new LinkedHashMap<>(rows.get(i));
                    String billNo = generateBillNumber(rowData, seqMap);
                    rowData.put("提单编号", billNo);
                    rowData.put("供方简称", firstFourChars(emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位"))));
                    rowData.put("需方简称", firstFourChars(emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"))));

                    try {
                        String outputFileName = resolveFileName(rowData, config.getOutputFileNamePattern());
                        outputFileName = ensureUniqueFileName(outputFileName, usedFileNames, i + 1);
                        usedFileNames.add(outputFileName);
                        Path outputFile = sheetOutputDir.resolve(outputFileName);
                        fillAndSave(templateFile, rowData, outputFile, ctx, i);
                        log.info("[sheet {}] [{}/{}] 已生成: {}", sheetName, i + 1, rows.size(), outputFileName);
                        successCount++;
                    } catch (Exception e) {
                        log.error("[sheet {}] 第 {} 行提单生成失败: {}", sheetName, i + 2, e.getMessage(), e);
                    }
                }

                log.info("sheet[{}] 提单批量生成完成，成功 {}/{} 个，输出目录: {}", sheetName, successCount, rows.size(), sheetOutputDir);
                totalSuccess += successCount;
            }
        }

        log.info("提单批量生成完成，成功 {} 个，输出根目录: {}", totalSuccess, outputRoot);
        return totalSuccess;
    }

    /**
     * 预计算卡号分配：按行顺序模拟整条链路的收发。
     *
     * 规则：
     * - 每家公司作为需方收货时，从「上游供方的池」或外部创建池生成卡号，并加入自己的池；
     * - 每家公司作为供方发货时，一律从自己的池中扣减（保证“收多少发多少”可追溯）；
     * - 行内合计始终等于数据表原始数量；同一家公司同一规格下的卡总量尽量保持互不重复。
     */
    private CardAllocationContext buildCardAllocationContext(List<Map<String, String>> rows) {
        log.info("[卡号分配] 开始 buildCardAllocationContext 总行数={}", rows.size());
        int min = config.getSplitMin();
        int max = config.getSplitMax();
        int maxSubCount = config.getSplitMaxSubCount();

        // 每个 (公司, 规格) 的卡池：key = 公司|规格
        Map<String, List<CardState>> poolsByCompany = new HashMap<>();
        // 每个 (公司, 规格) 已使用过的“单卡数量”，用于避免重复数量
        Map<String, Set<Integer>> usedAmountsByCompany = new HashMap<>();

        List<List<CardQty>> rowAllocations = new ArrayList<>(rows.size());
        for (int i = 0; i < rows.size(); i++) {
            rowAllocations.add(null);
        }

        for (int i = 0; i < rows.size(); i++) {
            Map<String, String> row = rows.get(i);
            String gongFang = emptyToBlank(getColumnValue(row, "供方/发货单位", "供方／发货单位"));
            String xuFang = emptyToBlank(getColumnValue(row, "需方/提货单位", "需方／提货单位"));
            String specModel = emptyToBlank(getColumnValue(row, "规格型号"));
            int qty = (int) Math.round(parseQuantity(row));
            if (qty <= 0) continue;

            String gongKey = !gongFang.isEmpty() ? gongFang + "|" + specModel : null;
            String xuKey = !xuFang.isEmpty() ? xuFang + "|" + specModel : null;

            List<CardQty> allocForRow = null;

            // 情形 1：供方有现成池（内部传递）：从供方池扣减，同时把同样的卡+数量记入需方池
            if (gongKey != null && poolsByCompany.containsKey(gongKey)) {
                List<CardState> fromPool = poolsByCompany.get(gongKey);
                allocForRow = allocateFromPool(fromPool, qty);
                if (allocForRow == null || allocForRow.isEmpty()) {
                    log.error("[卡号分配] 发货行 行{} 供方=[{}] qty={} 从池 [{}] 分配失败（库存不足或组合失败），跳过该行分配",
                            i + 2, gongFang, qty, gongKey);
                    continue;
                }
                // 将同样的卡号和数量加入需方池，实现卡号沿链路传递
                if (xuKey != null) {
                    List<CardState> toPool = poolsByCompany.computeIfAbsent(xuKey, k -> new ArrayList<>());
                    Set<Integer> usedSet = usedAmountsByCompany.computeIfAbsent(xuKey, k -> new HashSet<>());
                    for (CardQty cq : allocForRow) {
                        CardState target = null;
                        for (CardState c : toPool) {
                            if (c.cardNo.equals(cq.cardNo)) {
                                target = c;
                                break;
                            }
                        }
                        if (target == null) {
                            target = new CardState(cq.cardNo, 0);
                            toPool.add(target);
                        }
                        target.currentTotal += cq.qty;
                        usedSet.add(target.currentTotal);
                    }
                }
                StringBuilder sb = new StringBuilder();
                for (CardQty cq : allocForRow) sb.append(cq.cardNo).append(":").append(cq.qty).append(" ");
                log.info("[卡号分配] 发货行 行{} 供方=[{}] 需方=[{}] 规格=[{}] qty={} -> 从池[{}] 分配并转入 [{}]: {}",
                        i + 2, gongFang, xuFang, specModel, qty, gongKey, xuKey, sb.toString());
            } else if (xuKey != null) {
                // 情形 2：外部供货给某需方（或无供方信息）：为需方创建/扩展卡池
                List<CardState> pool = poolsByCompany.computeIfAbsent(xuKey, k -> new ArrayList<>());
                Set<Integer> usedSet = usedAmountsByCompany.computeIfAbsent(xuKey, k -> new HashSet<>());

                // 将现有卡总量加入 usedSet，避免新拆出的卡总量与已有卡冲突
                for (CardState c : pool) {
                    usedSet.add(c.currentTotal);
                }

                allocForRow = new ArrayList<>();
                int remainder = qty;

                // 先尝试对已有未满卡续收（小卡可续收）
                for (CardState card : pool) {
                    if (remainder <= 0) break;
                    if (card.currentTotal >= max) continue;
                    int stopFillAt = Math.max(0, max - STOP_FILL_GAP);
                    if (card.currentTotal >= stopFillAt) continue;
                    int room = max - card.currentTotal;
                    if (room <= 0) continue;
                    int bestAdd = 0;
                    for (int a = Math.min(remainder, room); a >= 1; a--) {
                        if (usedSet.contains(a)) continue;
                        int newTotal = card.currentTotal + a;
                        boolean conflict = false;
                        for (CardState other : pool) {
                            if (other == card) continue;
                            if (other.currentTotal == newTotal) {
                                conflict = true;
                                break;
                            }
                        }
                        if (!conflict) {
                            bestAdd = a;
                            break;
                        }
                    }
                    if (bestAdd <= 0) continue;
                    usedSet.add(bestAdd);
                    card.currentTotal += bestAdd;
                    allocForRow.add(new CardQty(card.cardNo, bestAdd));
                    remainder -= bestAdd;
                }

                // 剩余部分用 splitExcluding 创建新卡
                if (remainder > 0) {
                    List<Integer> amounts = OrderSplitUtil.splitExcluding(remainder, min, max, maxSubCount, usedSet);
                    if (amounts == null) {
                        amounts = new ArrayList<>();
                        amounts.add(remainder);
                        if (usedSet.contains(remainder)) {
                            log.warn("第 {} 行 (需方|规格={}) 数量 {} 无法在排除已用数量下拆分，该行使用单卡并分配新卡号",
                                    i + 2, xuKey, remainder);
                        }
                        usedSet.add(remainder);
                    } else {
                        usedSet.addAll(amounts);
                    }
                    int cardSeq = pool.size() + 1;
                    for (int a : amounts) {
                        String cardNo = specModel + "-" + cardSeq++;
                        CardState cs = new CardState(cardNo, a);
                        pool.add(cs);
                        allocForRow.add(new CardQty(cardNo, a));
                    }
                }

                // 行内合计兜底校正：保证本行分配合计严格等于原始数量
                int allocated = 0;
                for (CardQty cq : allocForRow) allocated += cq.qty;
                if (allocated != qty && !allocForRow.isEmpty()) {
                    int diff = qty - allocated;
                    CardQty last = allocForRow.get(allocForRow.size() - 1);
                    int newQty = last.qty + diff;
                    if (newQty > 0) {
                        allocForRow.set(allocForRow.size() - 1, new CardQty(last.cardNo, newQty));
                        for (CardState card : pool) {
                            if (card.cardNo.equals(last.cardNo)) {
                                card.currentTotal += diff;
                                usedSet.add(card.currentTotal);
                                break;
                            }
                        }
                        log.warn("[卡号分配] 行{} 需方|规格={} 原始数量={} 分配合计={} 存在差异，已对卡号 {} 调整 {}，修正后合计={}",
                                i + 2, xuKey, qty, allocated, last.cardNo, diff, qty);
                    } else {
                        log.error("[卡号分配] 行{} 需方|规格={} 原始数量={} 分配合计={} 且无法安全调整最后一条明细，保留原分配结果，请人工检查",
                                i + 2, xuKey, qty, allocated);
                    }
                }

                StringBuilder sb = new StringBuilder();
                for (CardQty cq : allocForRow) sb.append(cq.cardNo).append(":").append(cq.qty).append(" ");
                log.info("[卡号分配] 收货行 行{} 需方|规格={} qty={} -> 本行分配: {}", i + 2, xuKey, qty, sb.toString());
            }

            if (allocForRow != null && !allocForRow.isEmpty()) {
                rowAllocations.set(i, allocForRow);
            }
        }

        // 打印各公司卡池摘要，便于排查
        for (Map.Entry<String, List<CardState>> e : poolsByCompany.entrySet()) {
            String key = e.getKey();
            List<CardState> pool = e.getValue();
            StringBuilder poolSummary = new StringBuilder();
            for (CardState c : pool) poolSummary.append(c.cardNo).append("=").append(c.currentTotal).append(" ");
            log.info("[卡号分配] 最终卡池 key={} 共{}张卡: {}", key, pool.size(), poolSummary.toString());
        }

        return new CardAllocationContext(rowAllocations);
    }

    /**
     * 从供方卡池中分配出合计为 target 的 (卡号,数量)。
     * 单提单内：每张卡最多出现一次（可部分取），优先数量互不重复；若无法凑齐则放宽为允许数量重复以保证收发对齐。
     * 支持多笔发货：同一张卡可在不同提单中分批发出（池中扣减后剩余留给下一家）。
     */
    private List<CardQty> allocateFromPool(List<CardState> pool, int target) {
        if (target <= 0 || pool.isEmpty()) return Collections.emptyList();
        int total = 0;
        for (CardState c : pool) total += c.currentTotal;
        if (total < target) return null;

        // 特例：本单刚好等于当前池的总量，说明这批卡会在本单中全部发完。
        // 此时直接按卡池当前总量逐张开行，既保证每张卡“收多少发多少”，也天然满足数量互不重复（收货阶段已保证卡总量尽量唯一）。
        if (total == target) {
            List<CardQty> all = new ArrayList<>();
            for (CardState c : pool) {
                if (c.currentTotal <= 0) continue;
                all.add(new CardQty(c.cardNo, c.currentTotal));
                c.currentTotal = 0;
            }
            return all;
        }

        // 先尝试严格模式：数量互不重复
        List<CardQty> result = allocateFromPoolStrict(pool, target);
        if (result != null) return result;

        // 严格模式凑不齐时，用宽松模式：允许数量重复，保证池总量够时一定能凑出 target（收发对齐）
        log.debug("[卡号分配] 严格模式无法凑齐 target={}，改用宽松模式从池分配", target);
        return allocateFromPoolRelaxed(pool, target);
    }

    /** 严格模式：每张卡最多取一次，取的数量全局互不重复。 */
    private List<CardQty> allocateFromPoolStrict(List<CardState> pool, int target) {
        List<CardState> copy = new ArrayList<>();
        for (CardState c : pool) copy.add(new CardState(c.cardNo, c.currentTotal));

        List<CardQty> result = new ArrayList<>();
        Set<Integer> usedAmounts = new HashSet<>();
        Set<String> usedCardNos = new HashSet<>();
        int sum = 0;

        while (sum < target) {
            int need = target - sum;
            if (need <= 0) break;
            int chosenAmt = 0;
            CardState chosenCard = null;
            for (CardState card : copy) {
                if (card.currentTotal <= 0 || usedCardNos.contains(card.cardNo)) continue;
                int ideal = Math.min(need, card.currentTotal);
                if (ideal >= 1 && !usedAmounts.contains(ideal)) {
                    chosenCard = card;
                    chosenAmt = ideal;
                    break;
                }
            }
            if (chosenCard == null) {
                for (CardState card : copy) {
                    if (card.currentTotal <= 0 || usedCardNos.contains(card.cardNo)) continue;
                    int amt = findDistinctInRange(need, card.currentTotal, usedAmounts);
                    if (amt >= 1 && (chosenCard == null || amt > chosenAmt)) {
                        chosenCard = card;
                        chosenAmt = amt;
                    }
                }
            }
            if (chosenCard == null || chosenAmt < 1) break;
            result.add(new CardQty(chosenCard.cardNo, chosenAmt));
            usedAmounts.add(chosenAmt);
            usedCardNos.add(chosenCard.cardNo);
            chosenCard.currentTotal -= chosenAmt;
            sum += chosenAmt;
        }

        if (sum != target) return null;
        // 成功时回写扣减到原 pool（调用方依赖池被扣减）
        for (CardState c : copy) {
            for (CardState orig : pool) {
                if (orig.cardNo.equals(c.cardNo)) {
                    orig.currentTotal = c.currentTotal;
                    break;
                }
            }
        }
        return result;
    }

    /** 宽松模式：每张卡最多取一次，数量可重复，保证凑齐 target。 */
    private List<CardQty> allocateFromPoolRelaxed(List<CardState> pool, int target) {
        List<CardQty> result = new ArrayList<>();
        int sum = 0;
        for (CardState card : pool) {
            if (sum >= target) break;
            int need = target - sum;
            if (card.currentTotal <= 0) continue;
            int take = Math.min(need, card.currentTotal);
            result.add(new CardQty(card.cardNo, take));
            card.currentTotal -= take;
            sum += take;
        }
        if (sum != target) return null;
        return result;
    }

    /** 在 [1, min(need, cap)] 内取一个不在 used 中的值，优先接近 need。 */
    private static int findDistinctInRange(int need, int cap, Set<Integer> used) {
        int best = Math.min(need, cap);
        if (best < 1) return 0;
        if (!used.contains(best)) return best;
        for (int a = best + 1; a <= cap; a++) {
            if (!used.contains(a)) return a;
        }
        for (int a = best - 1; a >= 1; a--) {
            if (!used.contains(a)) return a;
        }
        return 0;
    }

    /** 卡池中一张卡的状态：卡号 + 当前总数量（可被续收直到 max） */
    private static class CardState {
        final String cardNo;
        int currentTotal;

        CardState(String cardNo, int currentTotal) {
            this.cardNo = cardNo;
            this.currentTotal = currentTotal;
        }
    }

    private static class CardQty {
        final String cardNo;
        final int qty;

        CardQty(String cardNo, int qty) {
            this.cardNo = cardNo;
            this.qty = qty;
        }
    }

    private static class CardAllocationContext {
        final List<List<CardQty>> rowAllocations;

        CardAllocationContext(List<List<CardQty>> rowAllocations) {
            this.rowAllocations = rowAllocations;
        }

        List<CardQty> getAllocation(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= rowAllocations.size()) return null;
            return rowAllocations.get(rowIndex);
        }
    }

    /**
     * 提单编号：供方/发货单位前4字拼音-需方/提货单位前4字拼音-签发时间-001
     */
    private String generateBillNumber(Map<String, String> rowData, Map<String, Integer> seqMap) {
        String gongFang = firstFourChars(emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位")));
        String xuFang = firstFourChars(emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位")));
        String dateStr = emptyToBlank(getColumnValue(rowData, "签发时间"));
        String gongInitials = toPinyinInitials(gongFang).toUpperCase();
        String xuInitials = toPinyinInitials(xuFang).toUpperCase();
        if (gongInitials.isEmpty()) gongInitials = "X";
        if (xuInitials.isEmpty()) xuInitials = "X";
        String yyyyMMdd = formatDate(dateStr);
        String groupKey = gongFang + "|" + xuFang + "|" + yyyyMMdd;
        int seq = seqMap.merge(groupKey, 1, Integer::sum);
        return gongInitials + "-" + xuInitials + "-" + yyyyMMdd + "-" + String.format("%03d", seq);
    }

    /** 按多个可能的列名取值，兼容全角/半角斜杠，以及简写表头（如“供方”“需方”）等 */
    private String getColumnValue(Map<String, String> rowData, String... possibleKeys) {
        // 扩展候选列名：在原始 possibleKeys 基础上，为“供方/发货单位”增加“供方”，为“需方/提货单位”增加“需方”
        List<String> candidates = new ArrayList<>();
        for (String pk : possibleKeys) {
            if (pk == null) continue;
            candidates.add(pk);
            if (pk.contains("供方")) candidates.add("供方");
            if (pk.contains("需方")) candidates.add("需方");
        }
        // 先按精确列名匹配
        for (String key : candidates) {
            String v = rowData.get(key);
            if (v != null && !v.trim().isEmpty()) return v;
        }
        // 再按“归一化后的列名”模糊匹配（兼容全角/半角等差异）
        for (Map.Entry<String, String> e : rowData.entrySet()) {
            String k = e.getKey();
            if (k == null) continue;
            String normK = normalizeKey(k);
            for (String pk : candidates) {
                if (normalizeKey(pk).equals(normK)) return e.getValue();
            }
        }
        return rowData.get(possibleKeys[0]);
    }

    private String normalizeKey(String key) {
        if (key == null) return "";
        return key.trim().replace("／", "/").replace("\u00A0", " ");
    }

    private String firstFourChars(String s) {
        if (s == null) return "";
        s = s.trim();
        int len = 0;
        for (int i = 0; i < s.length() && len < 4; i++) {
            len++;
        }
        return s.substring(0, Math.min(len, s.length()));
    }

    private String emptyToBlank(String s) {
        return s == null ? "" : s.trim();
    }

    private String formatDate(String dateStr) {
        if (dateStr == null || dateStr.trim().isEmpty()) {
            return LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE);
        }
        String s = dateStr.trim().replace(" ", "").replace("/", "-");
        DateTimeFormatter[] formatters = {
            DateTimeFormatter.ofPattern("yyyy-M-d"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ISO_LOCAL_DATE
        };
        for (DateTimeFormatter f : formatters) {
            try {
                return LocalDate.parse(s, f).format(DateTimeFormatter.BASIC_ISO_DATE);
            } catch (DateTimeParseException ignored) {}
        }
        return LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE);
    }

    private String toPinyinInitials(String str) {
        if (str == null || str.isEmpty()) return "";
        HanyuPinyinOutputFormat format = new HanyuPinyinOutputFormat();
        format.setToneType(HanyuPinyinToneType.WITHOUT_TONE);
        StringBuilder sb = new StringBuilder();
        for (char c : str.toCharArray()) {
            if (Character.toString(c).matches("[\\u4e00-\\u9fa5]")) {
                try {
                    String[] py = PinyinHelper.toHanyuPinyinStringArray(c, format);
                    if (py != null && py.length > 0) sb.append(py[0].charAt(0));
                } catch (Exception ignored) {}
            } else if (Character.isLetterOrDigit(c)) {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    /**
     * 单行独立拆分（回退逻辑）：将数量拆分为若干份，每份在 [splitMin, splitMax] 内互不重复
     */
    private List<Double> splitQuantity(double total) {
        if (total <= 0) return new ArrayList<>();
        int totalInt = (int) Math.round(total);
        int min = config.getSplitMin();
        int max = config.getSplitMax();
        int maxSubCount = config.getSplitMaxSubCount();
        if (totalInt < min) {
            return Collections.singletonList((double) totalInt);
        }
        try {
            List<Integer> parts = OrderSplitUtil.split(totalInt, min, max, maxSubCount);
            List<Double> result = new ArrayList<>(parts.size());
            for (Integer p : parts) {
                result.add(p.doubleValue());
            }
            return result;
        } catch (IllegalArgumentException e) {
            return Collections.singletonList((double) totalInt);
        }
    }

    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        PushbackInputStream pis = new PushbackInputStream(is, 8);
        byte[] header = new byte[4];
        int n = pis.read(header);
        if (n > 0) pis.unread(header, 0, n);
        boolean isXlsx = n >= 2 && header[0] == 0x50 && header[1] == 0x4B;
        return isXlsx ? new XSSFWorkbook(pis) : new HSSFWorkbook(pis);
    }

    private List<Map<String, String>> readDataSheet(Sheet sheet) {
        List<Map<String, String>> rows = new ArrayList<>();
        if (sheet == null || sheet.getPhysicalNumberOfRows() < 2) return rows;
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return rows;
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(getCellStringValue(cell, null));
        }
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Map<String, String> rowData = new LinkedHashMap<>();
            for (int c = 0; c < headers.size(); c++) {
                String header = headers.get(c);
                if (header == null || header.trim().isEmpty()) header = "col" + c;
                Cell cell = row.getCell(c);
                rowData.put(header.trim(), cell != null ? getCellStringValue(cell, header) : "");
            }
            // 跳过空行：供方和需方都为空，或数量不大于 0 的行不参与提单生成，避免生成大量空表
            String gongFang = emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位"));
            String xuFang = emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"));
            double qty = parseQuantity(rowData);
            if ((gongFang.isEmpty() && xuFang.isEmpty()) || qty <= 0) continue;
            rows.add(rowData);
        }
        return rows;
    }

    private String getCellStringValue(Cell cell, String columnHeader) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().format(DateTimeFormatter.ofPattern("yyyy/M/d"));
                }
                return formatNumeric(cell.getNumericCellValue(), -1);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return formatNumeric(cell.getNumericCellValue(), -1);
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e2) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    private static String formatNumeric(double num, int decimals) {
        if (Double.isNaN(num) || Double.isInfinite(num)) return String.valueOf(num);
        if (decimals == 0 || (decimals < 0 && num == Math.floor(num) && Math.abs(num) < 1e15)) {
            return String.valueOf((long) num);
        }
        int scale = decimals >= 0 ? decimals : 6;
        BigDecimal bd = BigDecimal.valueOf(num).setScale(scale, RoundingMode.HALF_UP);
        return bd.stripTrailingZeros().toPlainString();
    }

    private void fillAndSave(Path templateFile, Map<String, String> rowData, Path outputFile,
                             CardAllocationContext ctx, int rowIndex) throws IOException {
        try (InputStream is = Files.newInputStream(templateFile)) {
            Workbook workbook = openWorkbook(is, templateFile);

            replacePlaceholders(workbook, rowData);
            fillDetailTable(workbook, rowData, ctx, rowIndex);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
            if (workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook) {
                ((org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook).setForceFormulaRecalculation(true);
            } else if (workbook instanceof HSSFWorkbook) {
                ((HSSFWorkbook) workbook).setForceFormulaRecalculation(true);
            }

            try (OutputStream os = Files.newOutputStream(outputFile)) {
                workbook.write(os);
            }
            workbook.close();
        }
    }

    private Map<String, String> buildReplacementMap(Map<String, String> rowData) {
        Map<String, String> map = new HashMap<>(rowData);
        for (Map.Entry<String, String> e : new HashMap<>(rowData).entrySet()) {
            String normKey = normalizeKey(e.getKey());
            if (!normKey.equals(e.getKey())) map.put(normKey, e.getValue());
        }
        String gongFang = emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位"));
        String xuFang = emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"));
        String issueDate = emptyToBlank(getColumnValue(rowData, "签发时间"));
        map.put("供方/发货单位", gongFang);
        map.put("供方／发货单位", gongFang);
        map.put("需方/提货单位", xuFang);
        map.put("需方／提货单位", xuFang);
        map.put("签发时间", issueDate);
        return map;
    }

    private void replacePlaceholders(Workbook workbook, Map<String, String> rowData) {
        Map<String, String> replaceMap = buildReplacementMap(rowData);
        String prefix = config.getPlaceholderPrefix();
        String suffix = config.getPlaceholderSuffix();
        Pattern pattern = Pattern.compile(Pattern.quote(prefix) + "([^" + Pattern.quote(suffix) + "]+)" + Pattern.quote(suffix));

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            for (Row row : sheet) {
                if (row == null) continue;
                for (Cell cell : row) {
                    if (cell == null) continue;
                    if (cell.getCellType() == CellType.STRING) {
                        String value = cell.getStringCellValue();
                        String replaced = replaceInString(value, replaceMap, pattern, prefix, suffix);
                        if (!value.equals(replaced)) {
                            cell.setCellValue(replaced);
                        }
                    } else if (cell.getCellType() == CellType.FORMULA) {
                        try {
                            String formula = cell.getCellFormula();
                            if (pattern.matcher(formula).find()) {
                                String replaced = replaceInString(formula, replaceMap, pattern, prefix, suffix);
                                if (!formula.equals(replaced)) {
                                    cell.setCellFormula(replaced);
                                }
                            }
                        } catch (Exception ignored) {}
                    }
                }
            }
        }
    }

    private String replaceInString(String str, Map<String, String> replaceMap, Pattern pattern, String prefix, String suffix) {
        if (str == null) return "";
        Matcher matcher = pattern.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1).trim();
            String replacement = replaceMap.getOrDefault(key, replaceMap.getOrDefault(normalizeKey(key), ""));
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    /**
     * 查找表头行并填充数据行。优先使用预计算的卡号分配（保证跨单卡号对应），否则回退到单行独立拆分。
     */
    private void fillDetailTable(Workbook workbook, Map<String, String> rowData,
                                 CardAllocationContext ctx, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(0);
        int headerRowIdx = findTableHeaderRow(sheet);
        if (headerRowIdx < 0) {
            log.warn("未找到表头行（产品名、规格型号、卡号、数量（吨）、件数、备注），跳过表格填充");
            return;
        }

        int[] colIndices = findTableColumnIndices(sheet.getRow(headerRowIdx));
        if (colIndices == null) return;

        String productName = emptyToBlank(getColumnValue(rowData, "产品名"));
        String specModel = emptyToBlank(getColumnValue(rowData, "规格型号"));
        String xuFang = emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"));

        List<CardQty> allocation = ctx != null ? ctx.getAllocation(rowIndex) : null;
        if (allocation != null && !allocation.isEmpty()) {
            // 为了可读性，按卡号尾号从小到大排序，例如 52518-1, 52518-2, ...
            allocation.sort((a, b) -> {
                String ca = a.cardNo;
                String cb = b.cardNo;
                int ia = ca.lastIndexOf('-');
                int ib = cb.lastIndexOf('-');
                String sa = ia >= 0 ? ca.substring(ia + 1) : ca;
                String sb = ib >= 0 ? cb.substring(ib + 1) : cb;
                try {
                    int na = Integer.parseInt(sa);
                    int nb = Integer.parseInt(sb);
                    return Integer.compare(na, nb);
                } catch (NumberFormatException e) {
                    return ca.compareTo(cb);
                }
            });
        }
        List<Double> qtysForFallback = null;

        int firstDataRowIdx = headerRowIdx + 1;
        int templateLastRowIdx = firstDataRowIdx + 28;
        int dataRowCount;

        if (allocation == null || allocation.isEmpty()) {
            qtysForFallback = splitQuantity(parseQuantity(rowData));
            dataRowCount = qtysForFallback.isEmpty() ? 0 : qtysForFallback.size();
        } else {
            dataRowCount = allocation.size();
        }

        if (dataRowCount == 0) {
            writeTableRow(ensureRow(sheet, firstDataRowIdx), colIndices, productName, specModel,
                    specModel + "-1", 0, 0, "请过户至" + xuFang + specModel + "-1");
            dataRowCount = 1;
        } else {
            Row styleSourceRow = sheet.getRow(firstDataRowIdx);
            for (int i = 0; i < dataRowCount; i++) {
                double qty;
                int pieceCount;
                String cardNo;
                String remark;
                if (allocation != null && !allocation.isEmpty()) {
                    CardQty cq = allocation.get(i);
                    qty = cq.qty;
                    pieceCount = cq.qty * PIECES_MULTIPLIER;
                    cardNo = cq.cardNo;
                    remark = "请过户至" + xuFang + cq.cardNo;
                } else {
                    // 回退路径：使用预先计算好的 qtysForFallback，避免每次循环重新随机拆分导致长度不一致
                    qty = qtysForFallback.get(i);
                    pieceCount = (int) Math.round(qty * PIECES_MULTIPLIER);
                    cardNo = specModel + "-" + (i + 1);
                    remark = "请过户至" + xuFang + cardNo;
                }
                int targetRowIdx = firstDataRowIdx + i;
                Row row;
                if (targetRowIdx <= templateLastRowIdx) {
                    row = ensureRow(sheet, targetRowIdx);
                } else {
                    row = insertRowWithStyle(sheet, targetRowIdx, styleSourceRow);
                }
                writeTableRow(row, colIndices, productName, specModel, cardNo, qty, pieceCount, remark);
            }
        }
        int lastDataRowIdx = firstDataRowIdx + dataRowCount - 1;
        updateSumFormulasToLastRow(sheet, lastDataRowIdx);
    }

    private int findTableHeaderRow(Sheet sheet) {
        for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 30); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell == null) continue;
                String val = getCellStringValue(cell, null);
                if (val != null && val.contains("产品名")) {
                    return r;
                }
            }
        }
        return -1;
    }

    private int[] findTableColumnIndices(Row headerRow) {
        if (headerRow == null) return null;
        Map<String, Integer> headerToCol = new HashMap<>();
        for (Cell cell : headerRow) {
            String val = getCellStringValue(cell, null);
            if (val == null) continue;
            val = val.trim();
            for (String h : TABLE_HEADERS) {
                if (val.contains(h) || h.contains(val)) {
                    headerToCol.putIfAbsent(h, cell.getColumnIndex());
                    break;
                }
            }
        }
        int[] indices = new int[TABLE_HEADERS.length];
        for (int i = 0; i < TABLE_HEADERS.length; i++) {
            Integer col = headerToCol.get(TABLE_HEADERS[i]);
            if (col == null) return null;
            indices[i] = col;
        }
        return indices;
    }

    private double parseQuantity(Map<String, String> rowData) {
        String v = getColumnValue(rowData, "数量", "数量（吨）");
        if (v == null || v.trim().isEmpty()) return 0;
        try {
            return Double.parseDouble(v.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private void writeTableRow(Row row, int[] colIndices, String productName, String specModel,
                              String cardNo, double qty, int pieceCount, String remark) {
        if (row == null) return;
        setCellValue(getCell(row, colIndices[0]), productName);
        setCellValue(getCell(row, colIndices[1]), specModel);
        setCellValue(getCell(row, colIndices[2]), cardNo);
        setCellValue(getCell(row, colIndices[3]), qty);
        setCellValue(getCell(row, colIndices[4]), pieceCount);
        setCellValue(getCell(row, colIndices[5]), remark);
    }

    private Row ensureRow(Sheet sheet, int rowIdx) {
        Row row = sheet.getRow(rowIdx);
        return row != null ? row : sheet.createRow(rowIdx);
    }

    private Row insertRowWithStyle(Sheet sheet, int insertAt, Row styleSourceRow) {
        sheet.shiftRows(insertAt, sheet.getLastRowNum(), 1, true, false);
        Row newRow = sheet.createRow(insertAt);
        if (styleSourceRow != null) {
            int srcRowNum = styleSourceRow.getRowNum();
            newRow.setHeight(styleSourceRow.getHeight());
            short lastCol = styleSourceRow.getLastCellNum();
            for (short c = 0; c < lastCol; c++) {
                Cell srcCell = styleSourceRow.getCell(c);
                if (srcCell != null && srcCell.getCellStyle() != null) {
                    Cell newCell = newRow.createCell(c);
                    newCell.setCellStyle(srcCell.getCellStyle());
                }
            }
            copyMergedRegionsToRow(sheet, srcRowNum, insertAt);
        }
        return newRow;
    }

    private void copyMergedRegionsToRow(Sheet sheet, int sourceRow, int newRow) {
        List<CellRangeAddress> toAdd = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.getFirstRow() <= sourceRow && sourceRow <= r.getLastRow()) {
                toAdd.add(new CellRangeAddress(newRow, newRow, r.getFirstColumn(), r.getLastColumn()));
            }
        }
        for (CellRangeAddress r : toAdd) {
            sheet.addMergedRegion(r);
        }
    }

    private void updateSumFormulasToLastRow(Sheet sheet, int lastDataRowIdx) {
        int lastRowExcel = lastDataRowIdx + 1;
        Pattern rangePattern = Pattern.compile("(SUM\\()([^)]+)(\\))");
        for (Row row : sheet) {
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell == null || cell.getCellType() != CellType.FORMULA) continue;
                try {
                    String formula = cell.getCellFormula();
                    if (!formula.toUpperCase().contains("SUM(")) continue;
                    Matcher m = rangePattern.matcher(formula);
                    StringBuffer sb = new StringBuffer();
                    while (m.find()) {
                        String range = m.group(2);
                        String updated = range.replaceAll(":([A-Z$]+)\\d+$", ":$1" + lastRowExcel);
                        m.appendReplacement(sb, Matcher.quoteReplacement(m.group(1) + updated + m.group(3)));
                    }
                    m.appendTail(sb);
                    String newFormula = sb.toString();
                    if (!newFormula.equals(formula)) {
                        cell.setCellFormula(newFormula);
                    }
                } catch (Exception e) {
                    log.debug("跳过公式更新: {}", e.getMessage());
                }
            }
        }
    }

    private Cell getCell(Row row, int colIdx) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) cell = row.createCell(colIdx);
        return cell;
    }

    private void setCellValue(Cell cell, String value) {
        if (cell == null) return;
        if (value == null) value = "";
        value = value.trim();
        try {
            cell.setCellValue(Double.parseDouble(value.replace(",", "")));
        } catch (NumberFormatException e) {
            cell.setCellValue(value);
        }
    }

    private void setCellValue(Cell cell, double value) {
        if (cell == null) return;
        cell.setCellValue(value);
    }

    private void setCellValue(Cell cell, int value) {
        if (cell == null) return;
        cell.setCellValue(value);
    }

    private String resolveFileName(Map<String, String> rowData, String pattern) {
        Map<String, String> replaceMap = buildReplacementMap(rowData);
        String result = pattern;
        for (Map.Entry<String, String> e : replaceMap.entrySet()) {
            result = result.replace(config.getPlaceholderPrefix() + e.getKey() + config.getPlaceholderSuffix(),
                    e.getValue() != null ? e.getValue() : "");
        }
        result = result.replaceAll(Pattern.quote(config.getPlaceholderPrefix()) + "[^" + Pattern.quote(config.getPlaceholderSuffix()) + "]*" + Pattern.quote(config.getPlaceholderSuffix()), "");
        return result.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    private String ensureUniqueFileName(String fileName, Set<String> usedFileNames, int rowIndex) {
        if (!usedFileNames.contains(fileName)) return fileName;
        int lastDot = fileName.lastIndexOf('.');
        String base = lastDot > 0 ? fileName.substring(0, lastDot) : fileName;
        String ext = lastDot > 0 ? fileName.substring(lastDot) : "";
        return base + "_" + rowIndex + ext;
    }
}
