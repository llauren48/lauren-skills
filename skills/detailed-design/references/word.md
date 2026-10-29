# Word 输出

assets/template.xml 冻结于 2026-10-29，已移除飞书 block ID。生成项目稿时删“模板使用说明”，替换标题与修订记录；不要沿用模板作者、日期作为项目事实。按项目裁剪不适用章节。

依赖 Python 3.10+、scripts/requirements.txt 中的 python-docx。有宿主文档运行库先发现并使用；没有时可在独立虚拟环境安装，不修改项目依赖。

```text
python scripts/export_docx.py design.xml --output design.docx
```

脚本默认拒绝覆盖，明确需要覆盖时使用 --overwrite。内容源是 UTF-8 XML，以 document 为根，支持 title、h1～h6、p、b、em、u、br、a href、table/thead/tbody/tr/th/td、ul/ol/li、pre/code、blockquote、hr 和 img。表格单元格内使用 p，支持 colspan，不支持 rowspan。

图片使用 img path="本地 PNG/JPEG 路径" caption="说明"；相对路径基于内容源所在目录。SVG／Mermaid 先用当前可用渲染器转为 PNG 并检查；飞书画板也先导出图片。脚本遇到 token 图片、画板或其他不支持的资源会报错，防止静默丢失内容。

脚本保留加粗、斜体、超链接、标题层级和网格表格，使用 A4 页面及中文字体。生成后用 python-docx 重新打开检查正文、表格、图片和链接；有 Word／LibreOffice 渲染能力时检查分页与溢出，没有时如实标记未做视觉检查，不为此安装大型办公软件。交付真实 docx 绝对路径链接。

## 固化样式

样式由 assets/word-style.json 控制，默认 A4、Arial／等线、正文11磅，标题26磅，一级至四级标题18／16／15／14磅，1.2倍行距及浅灰表格边框。参考文档未显式定义页边距，因此采用2.54厘米作为确定性默认值。export_docx.py 自动应用此配置；assets/template.docx 已同步。

已有 Word 仅调整样式时使用 scripts/style_docx.py SOURCE --output OUTPUT，保留原有文字、表格、链接与图片；原文件不覆盖。该脚本通过原生标题样式或段落 outline level 识别标题，另将文档首段作为文档标题；特殊封面应先核对结构。
