---
name: commercial-bid
display_name: 商务标制作
display_name_en: Commercial Bid Maker
description: 制作政府采购与招投标项目的商务标文件。读取招标文件后提取资格条件、评审办法、废标条款和文件格式，根据供应商信息自动填充投标函、开标一览表、资格声明、中小企业声明函、售后服务承诺书等商务文件，并生成可编辑的 Word 草稿。适用场景：用户需要制作商务标、标书商务文件，或根据招标文件自动填充供应商信息时。
description_zh: 根据招标文件自动生成商务标 Word 草稿：投标函、开标一览表、资格声明、中小企业声明函、售后服务承诺书等。
description_en: Generates commercial bid Word drafts (bid letter, bid opening table, qualification statement, SME declaration, after-sales commitment, etc.) from tender documents and supplier information.
examples_zh:
  - "读取这份招标文件，给出商务标框架和供应商信息采集清单"
  - "根据 config.json 生成投标函和开标一览表"
examples_en:
  - "Read this tender document and output the commercial bid framework"
  - "Generate the bid letter and bid opening table from config.json"
version: 1.0.0
author: 智德管理
---

# 商务标制作

目标：根据招标文件要求制作商务文件（商务标），并根据用户提供的投标供应商信息自动填充相关内容。先出框架，后续按步填充。

## 硬性规则

- 所有落款日期统一填开标日期（投标函、开标一览表、资格声明、各类声明函、承诺书等的日期字段；资格声明使用中文大写日期）。
- 文件格式和排版严格按招标文件“格式一至格式七”执行，不得自创格式；招标文件未给出格式的（如无重大违法记录声明）才允许自拟。
- 电子标须同时加盖供应商公章（电子签章）和法定代表人名章（电子签章）；作为评分依据的材料必须盖章。
- 报价只允许一个总价，不超过采购预算；投标函、开标一览表、报价明细表三者金额必须一致，大小写一致。
- 红线：信用中国、中国政府采购网、中国裁判文书网查询时间必须在招标公告发布之后；专门面向中小企业的项目必须有中小企业声明函；暗标（技术标）不得出现供应商名称、人员姓名等可识别标识。
- 禁止代填或伪造不实材料（中小企业数据、业绩、社保、信用截图等必须真实），不参与围标串标、资质挂靠。

## 工作流程

1. 读取招标文件（PDF 用 pymupdf/fitz 提取文本，扫描件需 OCR），提取：项目信息、资格条件、评审办法、废标条款、目录结构和格式模板。
2. 输出商务标框架和供应商信息采集清单，明确还缺哪些材料。
3. 让用户把项目与供应商信息填入配置 JSON（模板见 `scripts/config.example.json`）。
4. 生成可编辑的 Word 草稿：`python scripts/generate_docs.py --config <配置.json> --out <输出目录>`。
5. 交付前逐项核对签字盖章、报价勾稽、装订密封和递交要求，重点排查废标风险。

## 需要核对格式时

格式一至格式七的逐条内容与废标风险清单见 [references/格式模板.md](references/格式模板.md)，生成或复核时按该文件对照招标文件执行。

## 注意

- 配置 JSON 包含营业执照、账户、审计报告等敏感信息，不要提交到公开仓库；只提交 `config.example.json` 模板。
- 生成结果默认包含【待填】占位符，交付前必须逐项填实并加盖电子签章。
