# 商务标制作技能（commercial-bid）

一套用于制作政府采购/招投标“商务标”的 Codex 技能：读招标文件、按格式生成投标函、开标一览表、资格声明、中小企业声明函、售后服务承诺书等商务文件，并根据供应商信息自动填充，输出可编辑的 Word 草稿。

## 一键安装（Codex 用户）

先把本仓库推送到 GitHub，然后对方在本机执行：

```bash
python "$CODEX_HOME/skills/.system/skill-installer/scripts/install-skill-from-github.py" \
  --repo <你的GitHub用户名>/<仓库名> \
  --path commercial-bid
```

安装后技能会出现在下一轮对话中。也可以手动安装：把 `commercial-bid` 文件夹复制到 `~/.codex/skills/`（或 `$CODEX_HOME/skills/`）下即可。

## 使用方式

1. 复制 `commercial-bid/scripts/config.example.json` 为 `config.json`。
2. 按实际情况填写项目信息和供应商信息。
3. 运行：

```bash
python commercial-bid/scripts/generate_docs.py --config config.json --out ./输出目录
```

生成的文件按招标文件“格式一至格式七”排版，落款日期自动填开标日期。生成后需逐项核对并加盖供应商公章和法定代表人名章（电子签章）。

## 说明

- 本项目仅做格式与文字初稿生成，不保证替代专业审查；使用前请逐条对照招标文件。
- 请勿用于伪造资质、业绩、社保证明等不实材料，不参与围标串标。
