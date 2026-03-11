package com.tjexcel.batch.runner;

import com.tjexcel.batch.service.BillOfLadingService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

/**
 * 启动时执行提单批量生成（在合同生成之后执行）
 */
@Component
@Order(2)
@ConditionalOnProperty(name = "batch.auto-run-bill", havingValue = "true", matchIfMissing = false)
public class BillOfLadingRunner implements CommandLineRunner {

    private static final Logger log = LoggerFactory.getLogger(BillOfLadingRunner.class);

    private final BillOfLadingService billOfLadingService;

    public BillOfLadingRunner(BillOfLadingService billOfLadingService) {
        this.billOfLadingService = billOfLadingService;
    }

    @Override
    public void run(String... args) throws Exception {
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
}
