# 报关单信息提取工具

从海关报关单XML报文中提取关键信息，结合外汇管理局汇率计算本币金额，并导出到Excel。

## 功能

- 自动下载并配置 Microsoft Edge WebDriver
- 从中国外汇管理局获取指定月份的美元汇率
- 解析海关报关单XML报文
- 提取：报关单号、合同号、出口日期、商品名称、数量、单位、原币金额、汇率、本币金额
- 导出结果到 Excel 文件

## 环境依赖

```bash
pip install selenium requests beautifulsoup4 pandas openpyxl
```

## 使用方法

1. 准备报关单XML文件（解密后），放入项目根目录
2. 运行脚本：

```bash
python automated_xml_processor.py
```

3. 脚本会自动：
   - 检测并下载匹配的 Edge WebDriver
   - 获取报关单日期对应的汇率
   - 解析XML数据并计算本币金额
   - 生成 `报关单清单提取结果.xlsx`

## 项目结构

```
报关单信息提取/
├── src/                  # 源代码（如有需要可拆分）
├── data/                 # XML数据文件
├── output/               # 导出结果
├── docs/                 # 文档
├── automated_xml_processor.py  # 主程序
└── README.md
```

## 注意事项

- XML文件需要提前解密
- 首次运行会自动下载 WebDriver
- 汇率数据来源于中国外汇管理局
