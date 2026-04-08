package com.tjexcel.batch.runner;

import com.tjexcel.batch.config.ContractConfig;
import com.tjexcel.batch.service.BillOfLadingService;
import com.tjexcel.batch.service.ContractGeneratorService;
import com.tjexcel.batch.service.FundFlowService;
import com.tjexcel.batch.service.UpstreamDownstreamService;
import com.tjexcel.batch.util.ExcelInspector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Scanner;

/**
 * 启动菜单：1合同 2提单 3资金流 4上下游数据表
 */
@Component
@Order(0)
public class BatchMenuRunner implements CommandLineRunner {

    private static final Logger log = LoggerFactory.getLogger(BatchMenuRunner.class);

    private final ContractGeneratorService contractGeneratorService;
    private final BillOfLadingService billOfLadingService;
    private final FundFlowService fundFlowService;
    private final UpstreamDownstreamService upstreamDownstreamService;
    private final ContractConfig contractConfig;

    public BatchMenuRunner(ContractGeneratorService contractGeneratorService,
                           BillOfLadingService billOfLadingService,
                           FundFlowService fundFlowService,
                           UpstreamDownstreamService upstreamDownstreamService,
                           ContractConfig contractConfig) {
        this.contractGeneratorService = contractGeneratorService;
        this.billOfLadingService = billOfLadingService;
        this.fundFlowService = fundFlowService;
        this.upstreamDownstreamService = upstreamDownstreamService;
        this.contractConfig = contractConfig;
    }

    @Override
    public void run(String... args) throws Exception {
        if (Arrays.asList(args).contains("--inspect")) {
            runInspect(args);
            return;
        }

        System.out.println();
        System.out.println("====================== 批量生成 ======================");
        System.out.println("请选择要生成的文件：");
        System.out.println("  1 - 合同");
        System.out.println("  2 - 提单");
        System.out.println("  3 - 资金流");
        System.out.println("  4 - 上下游数据表");
        System.out.println("  0 - 退出");
        System.out.println("====================================================");
        System.out.print("请输入选项 (0-4): ");

        String input;
        try (Scanner scanner = new Scanner(System.in)) {
            input = scanner.nextLine();
        }

        if (input == null) input = "";
        input = input.trim();

        switch (input) {
            case "1":
                runContract();
                break;
            case "2":
                runBillOfLading();
                break;
            case "3":
                runFundFlow();
                break;
            case "4":
                runUpstreamDownstream();
                break;
            case "0":
                log.info("用户选择退出");
                System.exit(0);
                break;
            default:
                log.warn("无效选项: {}", input);
                System.out.println("无效选项，请输入 0、1、2、3 或 4");
                System.exit(1);
        }
    }

    private void runContract() {
        log.info("开始执行合同批量生成...");
        try {
            int count = contractGeneratorService.generate();
            log.info("合同批量生成完成，共生成 {} 个文件", count);
            System.exit(0);
        } catch (Exception e) {
            log.error("合同批量生成失败", e);
            System.exit(1);
        }
    }

    private void runBillOfLading() {
        log.info("开始执行提单批量生成...");
        try {
            int count = billOfLadingService.generate();
            log.info("提单批量生成完成，共生成 {} 个文件", count);
            System.exit(0);
        } catch (Exception e) {
            log.error("提单批量生成失败", e);
            System.exit(1);
        }
    }

    private void runFundFlow() {
        log.info("开始执行资金流生成...");
        try {
            int count = fundFlowService.generate();
            log.info("资金流生成完成，共生成 {} 个文件", count);
            System.exit(0);
        } catch (Exception e) {
            log.error("资金流生成失败", e);
            System.exit(1);
        }
    }

    private void runUpstreamDownstream() {
        log.info("开始执行上下游数据表生成...");
        try {
            int count = upstreamDownstreamService.generate();
            log.info("上下游数据表生成完成，共生成 {} 个文件", count);
            System.exit(0);
        } catch (Exception e) {
            log.error("上下游数据表生成失败", e);
            System.exit(1);
        }
    }

    private void runInspect(String[] args) throws Exception {
        Path dataPath = Paths.get(contractConfig.getDataPath()).toAbsolutePath();
        Path templatePath = Paths.get(contractConfig.getTemplatePath()).toAbsolutePath();
        Path filledPath = dataPath.getParent().resolve("合同-伊科东城-奥卓-XMYKAZ2026113.xls");
        if (!Files.exists(filledPath)) {
            filledPath = Paths.get("src/main/resources/合同-伊科东城-奥卓-XMYKAZ2026113.xls").toAbsolutePath();
        }
        if (!Files.exists(filledPath)) {
            log.warn("未找到已填充样本，将只分析数据和模板");
        }
        Path outputFile = dataPath.getParent().resolve("inspect-output.txt");
        try (PrintStream ps = new PrintStream(Files.newOutputStream(outputFile))) {
            ExcelInspector.run(dataPath, templatePath, Files.exists(filledPath) ? filledPath : templatePath, ps);
            log.info("分析结果已写入: {}", outputFile);
        }
        ExcelInspector.run(dataPath, templatePath, Files.exists(filledPath) ? filledPath : templatePath, System.out);
        System.exit(0);
    }
}
