# 合同批量生成工具

根据数据表批量填充模板，生成合同文件。

## 快速开始

```bash
# 1. 打包
mvn clean package -DskipTests

# 2. 运行（使用默认配置）
java -jar target/batch-0.0.1-SNAPSHOT.jar

# 3. 或通过命令行指定路径
java -jar target/batch-0.0.1-SNAPSHOT.jar ^
  --contract.data-path=D:/data/0126.xlsx ^
  --contract.template-path=D:/templates/塑料合同模板(2).xlsx ^
  --contract.output-dir=D:/output
```

## 配置说明

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| contract.data-path | 填写数据表路径（第一个 sheet） | src/main/resources/0126.xlsx |
| contract.template-path | 模板文件路径 | src/main/resources/塑料合同模板(2).xlsx |
| contract.output-dir | 输出目录 | output |
| contract.output-file-name-pattern | 输出文件名模式 | 合同-${买方}-${卖方}-${合同编号}.xls |

## 文件要求

1. **数据表**：第一个 sheet 的第一行为列名（表头），从第二行开始为数据
2. **模板**：在需要替换的单元格中使用 `${列名}` 占位符，列名需与数据表表头一致
3. **输出文件名**：使用 `output-file-name-pattern` 配置，其中的 `${列名}` 会被替换为对应行的值

## 示例

数据表 0126.xlsx 表头若有：买方、卖方、合同编号 等列，模板中写 `${买方}`、`${卖方}`、`${合同编号}`。

输出文件名 `合同-${买方}-${卖方}-${合同编号}.xls` 会生成如：`合同-伊科东城-奥卓-XMYKAZ2026113.xls`。
