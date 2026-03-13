package com.tjexcel.batch.util;

import java.util.*;

/**
 * 订单/数量拆单工具：将总数拆分为若干份，每份在 [min,max] 内且互不重复，总和等于总数。
 * 采用确定性算法，相同输入始终得到相同输出，保证进出货卡号一一对应。
 */
public class OrderSplitUtil {

    /** 默认目标单份值（用于计算份数） */
    private static final int TARGET = 400;

    /**
     * @param total       总数
     * @param min         每份最小值
     * @param max         每份最大值
     * @param maxSubCount 最大子单数量
     * @return 拆分后的列表（互不重复， deterministic）
     */
    public static List<Integer> split(int total, int min, int max, int maxSubCount) {
        if (total < min) {
            return Collections.singletonList(total);
        }
        int minCount = computeMinCountDistinct(total, min, max);
        int maxCount = computeMaxCountDistinct(total, min, max);
        if (minCount > maxSubCount || maxCount < minCount) {
            throw new IllegalArgumentException(
                    "在当前约束下无法拆分: 需 " + minCount + "~" + maxCount + " 份，最大允许 " + maxSubCount + " 份");
        }
        int parts = computePartsDeterministic(total, minCount, maxCount, maxSubCount);
        int[] arr = buildDistinctSumDeterministic(parts, total, min, max);
        List<Integer> list = new ArrayList<>(parts);
        for (int v : arr) list.add(v);
        return list;
    }

    /**
     * 拆分总数，且每个数量均不在 exclude 中（用于同一 (供方,需方,规格) 下多行之间数量全局不重复）。
     *
     * @param exclude 已使用的数量集合，可为 null 或空表示不排除
     * @return 拆分后的列表；若在约束下无法拆分则返回 null
     */
    public static List<Integer> splitExcluding(int total, int min, int max, int maxSubCount,
                                               Set<Integer> exclude) {
        if (total < min) {
            if (exclude != null && exclude.contains(total)) return null;
            return Collections.singletonList(total);
        }

        Set<Integer> forbidden = exclude != null ? new HashSet<>(exclude) : Collections.emptySet();
        int minCount = computeMinCountDistinct(total, min, max);
        int maxCount = computeMaxCountDistinct(total, min, max);

        if (minCount > maxSubCount || maxCount < minCount) {
            return null;
        }

        int upper = Math.min(maxCount, maxSubCount);
        for (int parts = minCount; parts <= upper; parts++) {
            int[] arr = buildDistinctSumWithExclude(parts, total, min, max, forbidden);
            if (arr != null) {
                List<Integer> list = new ArrayList<>(parts);
                for (int v : arr) list.add(v);
                return list;
            }
        }
        return null;
    }

    /**
     * 确定性计算份数：优先接近 total/TARGET，且在 [minCount, min(maxCount, maxSubCount)] 内。
     */
    private static int computePartsDeterministic(int total, int minCount, int maxCount, int maxSubCount) {
        int preferred = Math.max(1, total / TARGET);
        int upper = Math.min(maxCount, maxSubCount);
        if (preferred <= minCount) return minCount;
        if (preferred >= upper) return upper;
        return Math.min(preferred, upper);
    }

    /**
     * 确定性构造 n 个互不重复的数，在 [min,max] 内且和为 total。
     * 从 base = [min, min+1, ..., min+n-1] 开始，按固定顺序增减至目标和。
     */
    private static int[] buildDistinctSumDeterministic(int n, int total, int min, int max) {
        int[] arr = new int[n];
        for (int i = 0; i < n; i++) {
            arr[i] = min + i;
        }
        int sum = n * min + n * (n - 1) / 2;
        int diff = total - sum;

        while (diff > 0) {
            int idx = chooseAddIndex(arr, min, max);
            if (idx < 0) break;
            arr[idx]++;
            diff--;
        }
        while (diff < 0) {
            int idx = chooseSubIndex(arr, min, max);
            if (idx < 0) break;
            arr[idx]--;
            diff++;
        }
        return arr;
    }

    /** 确定性选择可加 1 的下标：从末尾往前找第一个满足 arr[i] < max 且 arr[i]+1 不与其他值冲突的。 */
    private static int chooseAddIndex(int[] arr, int min, int max) {
        for (int i = arr.length - 1; i >= 0; i--) {
            if (arr[i] >= max) continue;
            int next = arr[i] + 1;
            boolean ok = true;
            for (int j = 0; j < arr.length; j++) {
                if (j != i && arr[j] == next) {
                    ok = false;
                    break;
                }
            }
            if (ok) return i;
        }
        return -1;
    }

    /** 确定性选择可减 1 的下标：从开头往后找第一个满足 arr[i] > min 且 arr[i]-1 不与其他值冲突的。 */
    private static int chooseSubIndex(int[] arr, int min, int max) {
        for (int i = 0; i < arr.length; i++) {
            if (arr[i] <= min) continue;
            int next = arr[i] - 1;
            boolean ok = true;
            for (int j = 0; j < arr.length; j++) {
                if (j != i && arr[j] == next) {
                    ok = false;
                    break;
                }
            }
            if (ok) return i;
        }
        return -1;
    }

    /** 在 [min,max] 且不在 forbidden 中构造 n 个互不重复的数且和为 total；不可能时返回 null。 */
    private static int[] buildDistinctSumWithExclude(int n, int total, int min, int max,
                                                    Set<Integer> forbidden) {
        List<Integer> allowed = new ArrayList<>();
        for (int v = min; v <= max; v++) {
            if (!forbidden.contains(v)) allowed.add(v);
        }
        if (allowed.size() < n) return null;

        long minSum = 0;
        long maxSum = 0;
        for (int i = 0; i < n; i++) {
            minSum += allowed.get(i);
            maxSum += allowed.get(allowed.size() - 1 - i);
        }
        if (total < minSum || total > maxSum) return null;

        int[] arr = new int[n];
        for (int i = 0; i < n; i++) arr[i] = allowed.get(i);
        long sum = minSum;
        int diff = total - (int) sum;

        while (diff > 0) {
            int idx = chooseAddIndexAllowed(arr, allowed);
            if (idx < 0) return null;
            int cur = arr[idx];
            int pos = allowed.indexOf(cur);
            if (pos >= allowed.size() - 1) return null;
            int nextVal = allowed.get(pos + 1);
            boolean collision = false;
            for (int j = 0; j < arr.length; j++)
                if (j != idx && arr[j] == nextVal) { collision = true; break; }
            if (collision) return null;
            arr[idx] = nextVal;
            diff--;
        }
        while (diff < 0) {
            int idx = chooseSubIndexAllowed(arr, allowed);
            if (idx < 0) return null;
            int cur = arr[idx];
            int pos = allowed.indexOf(cur);
            if (pos <= 0) return null;
            int nextVal = allowed.get(pos - 1);
            boolean collision = false;
            for (int j = 0; j < arr.length; j++)
                if (j != idx && arr[j] == nextVal) { collision = true; break; }
            if (collision) return null;
            arr[idx] = nextVal;
            diff++;
        }
        return arr;
    }

    private static int chooseAddIndexAllowed(int[] arr, List<Integer> allowed) {
        for (int i = arr.length - 1; i >= 0; i--) {
            int pos = allowed.indexOf(arr[i]);
            if (pos < 0 || pos >= allowed.size() - 1) continue;
            int nextVal = allowed.get(pos + 1);
            boolean ok = true;
            for (int j = 0; j < arr.length; j++)
                if (j != i && arr[j] == nextVal) { ok = false; break; }
            if (ok) return i;
        }
        return -1;
    }

    private static int chooseSubIndexAllowed(int[] arr, List<Integer> allowed) {
        for (int i = 0; i < arr.length; i++) {
            int pos = allowed.indexOf(arr[i]);
            if (pos <= 0) continue;
            int nextVal = allowed.get(pos - 1);
            boolean ok = true;
            for (int j = 0; j < arr.length; j++)
                if (j != i && arr[j] == nextVal) { ok = false; break; }
            if (ok) return i;
        }
        return -1;
    }

    /** n 个 [min,max] 内互不相同的数，最小和为 n*min+n(n-1)/2。要求 total >= 该值，解得 n 的上界 */
    private static int computeMaxCountDistinct(int total, int min, int max) {
        int n = max - min + 1;
        for (int k = n; k >= 1; k--) {
            long minSum = (long) k * min + (long) k * (k - 1) / 2;
            if (minSum <= total) return k;
        }
        return 0;
    }

    /** n 个 [min,max] 内互不相同的数，最大和为 n*max-n(n-1)/2。要求 total <= 该值，解得 n 的下界 */
    private static int computeMinCountDistinct(int total, int min, int max) {
        int n = max - min + 1;
        for (int k = 1; k <= n; k++) {
            long maxSum = (long) k * max - (long) k * (k - 1) / 2;
            if (maxSum >= total) return k;
        }
        return n;
    }
}
