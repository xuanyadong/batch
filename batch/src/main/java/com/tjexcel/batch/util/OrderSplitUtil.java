package com.tjexcel.batch.util;

import java.util.*;
import java.util.concurrent.ThreadLocalRandom;

/**
 * 订单/数量拆单工具：将总数拆分为若干份，每份在 [min,max] 内且互不重复，总和等于总数。
 * 采用构造性算法保证 100% 成功。
 */
public class OrderSplitUtil {

    /**
     * @param total        总数
     * @param min          每份最小值
     * @param max          每份最大值
     * @param maxSubCount  最大子单数量
     * @return 拆分后的列表（互不重复）
     */
    public static List<Integer> split(int total, int min, int max, int maxSubCount) {
        if (total < min) {
            return Collections.singletonList(total);
        }

        int minCount = computeMinCountDistinct(total, min, max);
        int maxCount = computeMaxCountDistinct(total, min, max);

        if (minCount > maxSubCount || maxCount < minCount) {
            throw new IllegalArgumentException("在当前约束下无法拆分: 需 " + minCount + "~" + maxCount + " 份，最大允许 " + maxSubCount + " 份");
        }

        ThreadLocalRandom random = ThreadLocalRandom.current();
        int count = random.nextInt(minCount, Math.min(maxCount, maxSubCount) + 1);

        int[] arr = buildDistinctSum(count, total, min, max, random);

        List<Integer> list = new ArrayList<>();
        for (int v : arr) {
            list.add(v);
        }
        Collections.shuffle(list);

        return list;
    }

    /**
     * 构造 n 个互不重复的数，在 [min,max] 内且和为 total。从最小集合开始，逐个加 1 分配到各位置。
     */
    private static int[] buildDistinctSum(int n, int total, int min, int max, ThreadLocalRandom random) {
        int[] arr = new int[n];
        for (int i = 0; i < n; i++) {
            arr[i] = min + i;
        }
        int sum = n * min + n * (n - 1) / 2;
        int diff = total - sum;

        while (diff > 0) {
            List<Integer> canAdd = new ArrayList<>();
            for (int i = 0; i < n; i++) {
                if (arr[i] < max) {
                    int next = arr[i] + 1;
                    boolean collision = false;
                    for (int j = 0; j < n; j++) {
                        if (j != i && arr[j] == next) {
                            collision = true;
                            break;
                        }
                    }
                    if (!collision) {
                        canAdd.add(i);
                    }
                }
            }
            if (canAdd.isEmpty()) break;

            int idx = canAdd.get(random.nextInt(canAdd.size()));
            arr[idx]++;
            diff--;
        }

        while (diff < 0) {
            List<Integer> canSub = new ArrayList<>();
            for (int i = 0; i < n; i++) {
                if (arr[i] > min) {
                    int next = arr[i] - 1;
                    boolean collision = false;
                    for (int j = 0; j < n; j++) {
                        if (j != i && arr[j] == next) {
                            collision = true;
                            break;
                        }
                    }
                    if (!collision) {
                        canSub.add(i);
                    }
                }
            }
            if (canSub.isEmpty()) break;

            int idx = canSub.get(random.nextInt(canSub.size()));
            arr[idx]--;
            diff++;
        }

        return arr;
    }

    /** n 个 [min,max] 内互不相同的数，最小和为 n*min+n(n-1)/2。要求 total >= 该值，解得 n 的上界 */
    private static int computeMaxCountDistinct(int total, int min, int max) {
        int n = (max - min + 1);
        for (int k = n; k >= 1; k--) {
            long minSum = (long) k * min + (long) k * (k - 1) / 2;
            if (minSum <= total) return k;
        }
        return 0;
    }

    /** n 个 [min,max] 内互不相同的数，最大和为 n*max-n(n-1)/2。要求 total <= 该值，解得 n 的下界 */
    private static int computeMinCountDistinct(int total, int min, int max) {
        int n = (max - min + 1);
        for (int k = 1; k <= n; k++) {
            long maxSum = (long) k * max - (long) k * (k - 1) / 2;
            if (maxSum >= total) return k;
        }
        return n;
    }
}
