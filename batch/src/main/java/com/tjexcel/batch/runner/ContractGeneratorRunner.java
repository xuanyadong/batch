package com.tjexcel.batch.runner;

import com.tjexcel.batch.config.ContractConfig;
import com.tjexcel.batch.service.ContractGeneratorService;
import com.tjexcel.batch.util.ExcelInspector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

/**
 * 启动时执行合同批量生成（先于提单执行）
 */
@Component
@Order(1)
@ConditionalOnProperty(name = "batch.auto-run-contract", havingValue = "true", matchIfMissing = false)
public class ContractGeneratorRunner implements CommandLineRunner {

    private static final Logger log = LoggerFactory.getLogger(ContractGeneratorRunner.class);

    private final ContractGeneratorService generatorService;
    private final ContractConfig config;

    public ContractGeneratorRunner(ContractGeneratorService generatorService, ContractConfig config) {
        this.generatorService = generatorService;
        this.config = config;
    }

    @Override
    public void run(String... args) throws Exception {
        if (Arrays.asList(args).contains("--inspect")) {
            runInspect(args);
            return;
        }
        log.info("开始执行合同批量生成...");
        try {
            int count = generatorService.generate();
            log.info("合同批量生成完成，共生成 {} 个文件", count);
        } catch (Exception e) {
            log.error("合同批量生成失败", e);
            System.exit(1);
        }
    }

    private void runInspect(String[] args) throws Exception {
        Path dataPath = Paths.get(config.getDataPath()).toAbsolutePath();
        Path templatePath = Paths.get(config.getTemplatePath()).toAbsolutePath();
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
