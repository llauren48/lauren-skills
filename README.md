# Lauren Skills

## detailed-design：详细设计文档

根据需求、实际代码与官方资料编写技术详细设计。内置 XML／Word 模板，支持飞书文档或本地 Word 输出，覆盖架构、代码范围、数据适配、协议兼容、验收和测试。

### 安装与使用

将 [skills/detailed-design](skills/detailed-design) 文件夹复制到 `~/.codex/skills/`，或下载 [ZIP 包](packages/detailed-design.zip) 解压至技能目录。

```text
使用 $detailed-design，基于当前项目编写详细设计，输出为 Word。
```

未指定输出形式时会询问飞书或 Word。飞书输出需要可用且已授权的连接器或 lark-cli，Word 无需飞书。

### Word 导出

Python 3.10+，建议在独立虚拟环境安装依赖：

```bash
python -m pip install -r skills/detailed-design/scripts/requirements.txt
python skills/detailed-design/scripts/export_docx.py design.xml --output design.docx
```

格式见 [Word 指引](skills/detailed-design/references/word.md)。支持中文、富文本、表格、链接和 PNG/JPEG 图片，默认拒绝覆盖已有文件。

技能源码位于 `skills/`，ZIP 和 .skill 分发包位于 `packages/`。
