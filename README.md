# Lauren Skills

## detailed-design：详细设计文档

根据需求、实际代码与官方资料编写技术详细设计。内置 XML／Word 模板，支持飞书文档或本地 Word 输出，覆盖架构、代码范围、数据适配、协议兼容、验收和测试。

### 安装与使用

从 [skills/detailed-design](skills/detailed-design) 复制完整技能目录，或下载 [ZIP 包](packages/detailed-design.zip) 解压。不同 Agent 的安装位置见下文；当前 Codex 官方推荐 `.agents/skills/`，本项目创建时使用的 Codex Desktop 环境也支持 `~/.codex/skills/`，以实际客户端版本为准。

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

### Word 样式与已有文档重排

样式固化在 [`word-style.json`](skills/detailed-design/assets/word-style.json)：A4、Arial／等线、正文 11 磅、标题 26 磅，一级至四级标题 18／16／15／14 磅、1.2 倍行距和浅灰表格边框。默认页边距为 2.54 厘米。修改该配置后，新生成文档会自动使用对应样式；Word 模板同步提供在 `assets/template.docx`。

已有 Word 仅调整样式，使用独立脚本并输出新文件：

```bash
python skills/detailed-design/scripts/style_docx.py source.docx --output styled.docx
```

该脚本保留正文、表格、链接及图片，通过已有标题样式或段落大纲级别识别章节，并把首段设为文档标题；特殊封面需先检查结构。不会覆盖源文件，输出已存在时也默认拒绝覆盖。

技能源码位于 `skills/`，ZIP 和 .skill 分发包位于 `packages/`。

## 适配到不同 Coding Agent

本技能的可复用部分是 `SKILL.md`、`references/`、`assets/` 和 `scripts/`。迁移过程主要是**安装目录、调用入口和工具权限适配**，不需要为每个 Agent 重写详细设计模板。`agents/openai.yaml` 是 Codex 的界面元数据，不作为其他 Agent 的配置文件。

以下目录依据 2026-10-29 核对的官方说明，主要针对本地 Coding Agent。**目录可被发现不等于已完成各客户端实测**；本仓库已验证技能格式及 Word 脚本，其他客户端须执行下面的安装验证。

| Coding Agent | 项目级目录（相对目标项目根目录） | 个人级目录 | 调用方式 |
| --- | --- | --- | --- |
| Codex | `.agents/skills/detailed-design/` | `~/.agents/skills/detailed-design/` | `$detailed-design`，或技能选择器 |
| Claude Code | `.claude/skills/detailed-design/` | `~/.claude/skills/detailed-design/` | `/detailed-design`，或明确要求使用该技能 |
| Cursor | `.cursor/skills/detailed-design/` | `~/.cursor/skills/detailed-design/` | 在 Agent 中选择对应技能，或明确要求使用 `detailed-design` |
| GitHub Copilot（VS Code） | `.github/skills/detailed-design/` | `~/.copilot/skills/detailed-design/` | 在 Agent Chat 中明确指定技能；可用客户端入口以实际版本为准 |

### 用一句提示词安装

将下面这句话发给目标 Coding Agent：

> 请从 https://github.com/llauren48/lauren-skills 安装 `skills/detailed-design` 技能，根据你当前的平台选择项目级技能目录，完整保留模板、参考文档和脚本；已有同名技能时先比较差异再处理，安装后检查技能能否被发现，并告诉我调用方式。

请在目标项目中执行；如希望跨项目使用，将提示词中的“项目级”改为“个人级”。安装会写入技能目录，涉及额外权限时按当前 Agent 的提示处理，飞书授权仍需单独配置。

### 1. 复制完整目录

先克隆本仓库，再根据目标 Agent 选择一种目录。下面以**目标项目使用 Claude Code**为例，PowerShell 命令可将目录替换为表中的其他项目级位置：

```powershell
git clone https://github.com/llauren48/lauren-skills.git
# 将此路径替换为实际目标项目，不要直接使用占位路径
$projectRoot = 'D:/path/to/your-project'
$skillTarget = Join-Path $projectRoot '.claude/skills/detailed-design'
if (Test-Path -LiteralPath $skillTarget) {
    throw '目标技能已存在，请先比较版本，避免意外覆盖。'
}
New-Item -ItemType Directory -Force -Path (Split-Path $skillTarget) | Out-Null
Copy-Item -LiteralPath './lauren-skills/skills/detailed-design' -Destination $skillTarget -Recurse
```

目录最终应是 `<技能根目录>/detailed-design/SKILL.md`，而不是再套一层 `detailed-design/detailed-design/`。**不能只复制 SKILL.md**，否则会缺少固化模板和导出脚本。项目级安装适合团队共享，个人级安装适合跨项目复用；避免在多个被扫描目录中重复安装同名版本。

ZIP 是通用的目录分发包；`.skill` 是本仓库提供的打包文件，**不保证每个 Agent 都支持直接导入该扩展名**。跨 Agent 推荐使用源码目录或解压 ZIP。

### 2. 适配各 Agent 的执行能力

- **Codex**：保留 `agents/openai.yaml`。其他客户端无需转换该文件，也不应把其中的 `$detailed-design` 当作通用斜杠命令。
- **Claude Code**：复制到 `.claude/skills/` 后使用 `/detailed-design`。SKILL.md 中提及的 Codex 行为按当前 Claude Code 的工具执行，项目规则仍以目标项目为准。
- **Cursor**：在 Agent 模式中使用，确认允许读取目标项目、执行 Python 及写出文件。仅编辑器补全模式不能代替完整文档生成流程。
- **Copilot**：使用支持 Agent Skills 的 VS Code／Copilot 版本和 Agent 会话，确认工作区信任、终端和文件工具可用；其他 Copilot 客户端按各自支持情况检查，不套用 VS Code 结论。

遇到客户端不自动发现技能时，新开会话或刷新技能列表，检查安装路径、SKILL.md 的 name／description 和目录嵌套。暂不支持 Skills 的 Agent 可以按完整路径手动读取 SKILL.md 和引用文件，但这只是手动执行流程，不代表已完成原生技能注册。

### 3. 选择 Word 或飞书工具链

**Word 输出**只要求目标 Agent 能读写本地文件并运行 Python。建议在独立虚拟环境中安装 `scripts/requirements.txt`，然后按 [Word 指引](skills/detailed-design/references/word.md) 生成内容源并调用导出脚本。脚本和模板路径应相对**实际安装后的技能目录**定位，不假定终端当前目录就是技能目录。图示先渲染成本地 PNG/JPEG；没有渲染工具时明确说明，不插空白占位图。

**飞书输出**需要目标 Agent 所在环境提供已授权的飞书连接器，或安装并配置 `lark-cli`，并具备目标文档编辑权限。遵循 [飞书指引](skills/detailed-design/references/feishu.md) 读取最新内容、局部更新、回读核验。原环境中的 `lark-doc` 等辅助技能不会随本包自动安装；若不存在，使用该环境现有工具文档和本包指引。不复制账户登录状态、Token 或真实 API Key。

本技能不捆绑固定工具函数名或 MCP 服务配置。不同 Agent 的 MCP 配置文件、认证和终端权限应分别设置；**安装成功不代表飞书已授权，也不代表模型服务协议已经兼容**。远程／云端 Agent 还须在其实际执行环境准备技能和依赖，不能假定能读取本机个人技能目录。

### 4. 安装后的验证

1. 让 Agent 读取内置模板，说明准备采用的章节，检查模板与引用文件均可找到。
2. 用一个小项目生成 Word：核对真实 `.docx`、标题、中文、表格、加粗斜体、链接和图片；没有视觉渲染器时区分结构检查与视觉检查。
3. 需要飞书时，在获准使用的测试文档上执行一次“读取 → 局部写入 → 回读”，确认没有丢失其他内容。
4. 测试一次增量修改，确认不会覆盖用户手工修改，也不会把未执行测试写成通过。

### 官方参考

- [Codex：技能构建与发现目录](https://learn.chatgpt.com/docs/build-skills)
- [Claude Code：Skills](https://code.claude.com/docs/en/skills)
- [Cursor：Agent Skills](https://cursor.com/docs/skills)
- [GitHub Copilot / VS Code：Agent Skills](https://code.visualstudio.com/docs/agent-customization/agent-skills)
