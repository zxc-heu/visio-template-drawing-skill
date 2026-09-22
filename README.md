# Visio Template Drawing Skill

用中文告诉 Codex 要画什么，或提供 PNG 参考图、Visio 模板，即可生成原生可编辑的 `.vsdx`：形状、文字和连接线都可以继续在 Microsoft Visio 中修改。

**默认后台绘制，完成后关闭本任务的 Visio 实例。** 绘图过程不使用鼠标、键盘或系统剪贴板；你可以继续操作其他软件。只有明确要求时，才显示绘制过程、保留最终窗口或导出 PNG。

[下载稳定版 v1.0.0](https://github.com/zxc-heu/visio-template-drawing-skill/releases/tag/v1.0.0) · [使用方式](#usage) · [安装](#installation) · [命令行示例](#powershell) · [常见问题](#faq)

<a name="usage"></a>

## 日常使用：直接用中文提出要求

安装技能后，把参考图交给 Codex，复制下面的指令并补充自己的要求即可。**你不需要先写 JSON，也不需要记住脚本参数。** Codex 会分析输入、准备绘图计划、运行脚本并检查结果。

### 根据参考图绘制

```text
使用 visio-template-drawing，根据我上传的这张参考图生成可编辑 Visio。
保留图中的文字、结构、箭头方向和主要配色，后台绘制，输出到桌面。
```

### 根据文字绘制

```text
使用 visio-template-drawing，画一个“开始 → 输入数据 → 处理 → 输出结果 → 结束”的流程图。
中文使用宋体，节点浅蓝色，连接线带箭头，保存为桌面的“处理流程.vsdx”。
```

### 使用已有模板

```text
使用 visio-template-drawing，基于我提供的 Visio 模板绘制这张参考图。
优先复用模板中的形状和母版，另存为新的 VSDX，不覆盖原模板。
```

### 生成多页图册

```text
使用 visio-template-drawing，把这八张参考图绘制到同一个可编辑 VSDX 中。
每张参考图独立一页，按图1至图8排序，后台绘制，完成后关闭本任务窗口。
```

多页图册是在同一文档中按计划逐页绘制；当前不提供任意已有 VSDX 文件的页面合并或跨页连接导入。

### 选择显示方式

在请求末尾加一句即可；不加时默认后台绘制并关闭。

| 你的要求 | 对应参数 | 实际行为 |
| --- | --- | --- |
| 后台绘制，完成后关闭 | 默认 | 全程隐藏，保存后清理所属实例 |
| 让我看到绘制过程，完成后关闭 | `-Visible` | 显示过程，完成后关闭 |
| 后台绘制，完成后打开最终图并保留 | `-KeepOpen` | 全部成功后显示并保留最终文档 |
| 显示绘制过程，完成后保留最终图 | `-Visible -KeepOpen` | 显示过程，全部成功后保留最终文档 |
| 另外导出每一页的 PNG 预览 | `-PreviewDirectory` | 按页导出 PNG；VSDX 仍是可编辑源文件 |

即使选择保留窗口，任务失败时也会清理所属资源。可见模式可能改变焦点；希望持续操作其他软件时使用默认后台模式。绘图进行中不要手动编辑该任务正在生成的文档，成功交付后再编辑。

在支持显式技能选择的 Codex 界面中，也可以选择或引用 `visio-template-drawing`；CLI / IDE 扩展支持 `/skills` 或 `$visio-template-drawing`。技能安装与发现机制见 [OpenAI 官方说明](https://learn.chatgpt.com/docs/build-skills)。

<a name="installation"></a>

## 安装

### 环境要求

- Windows，本机已安装 Microsoft Visio。本项目在 Visio 16 上完成了实测。
- Windows PowerShell 5.1：实际 COM 绘图进程使用 STA 模式运行。也可以从 PowerShell 7 启动入口脚本。
- Python 3：用于复杂图的布局预检；只有截图裁剪工具需要额外安装 Pillow。
- 使用自然语言绘图时，需要能访问本机文件并执行本机脚本的 Codex 环境。本技能不会安装 Microsoft Visio。

### 方式一：让 Codex 安装

向 Codex 发送：

```text
使用 skill-installer，从以下仓库的 v1.0.0 版本安装技能：
https://github.com/zxc-heu/visio-template-drawing-skill
技能子目录：skill/visio-template-drawing
请先检查现有安装；如有本地改动，保留并说明差异，不直接覆盖。
```

### 方式二：下载后手动安装

1. 打开 [v1.0.0 发行版页面](https://github.com/zxc-heu/visio-template-drawing-skill/releases/tag/v1.0.0)，下载 **Source code (zip)** 并解压。
2. 找到解压目录中的 `skill/visio-template-drawing/`。需要安装的是这个子目录，不是整个仓库。
3. 将该目录放入你的 Codex 技能目录，确认最终路径直接包含 `visio-template-drawing/SKILL.md`，不要多套一层同名文件夹。

当前 [官方本地发现路径](https://learn.chatgpt.com/docs/build-skills) 为用户级 `~/.agents/skills/`，或项目级 `.agents/skills/`。如果你的环境已经从其他位置识别并使用此技能，例如维护环境中的 `~/.codex/skills/`，请继续更新原位置，避免重复安装同名技能。未识别新安装时可重启 Codex 再检查。

下面是**全新用户级安装**示例。在解压后的仓库根目录打开 PowerShell 执行；目标已存在时会停止，不覆盖旧文件。

```powershell
$source = (Resolve-Path '.\skill\visio-template-drawing').Path
$skillsDir = Join-Path $env:USERPROFILE '.agents\skills'
$destination = Join-Path $skillsDir 'visio-template-drawing'
if (Test-Path -LiteralPath $destination) {
  throw '目标技能已存在，请先备份并比较差异，再逐文件更新。'
}
New-Item -ItemType Directory -Path $skillsDir -Force | Out-Null
Copy-Item -LiteralPath $source -Destination $destination -Recurse
```

已经安装过时，先备份现有技能、核对差异，再更新需要替换的文件；保留自己增加的文件。GitHub 发布新版本不会自动覆盖你的本机安装。

<a name="powershell"></a>

## 自己运行 PowerShell（可选）

自然语言使用者可以跳过本节。以下命令从**解压后的仓库根目录**执行，不要求先把技能安装到某个固定位置。

### 先生成一个简单示例

```powershell
$skillRoot = (Resolve-Path '.\skill\visio-template-drawing').Path
$desktop = [Environment]::GetFolderPath('Desktop')
$runName = 'visio-demo-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
$output = Join-Path $desktop ($runName + '.vsdx')

& "$skillRoot\scripts\visio_draw_from_plan.ps1" `
  -PlanJson "$skillRoot\examples\case-08-background-flow.json" `
  -OutputVsdx $output
```

输出是桌面的一个可编辑流程图，保存后实例自动退出。在最后一行末尾继续加 `-KeepOpen`，就会在成功后显示并保留最终文档；加 `-Visible` 则显示绘制过程。注意：PowerShell 的反引号用于续行，后面不要留空格。

### 八页图册、可见模式和 PNG 预览

以下示例接着上面的变量运行，使用不同文件名。

```powershell
& "$skillRoot\scripts\visio_draw_from_plan.ps1" `
  -PlanJson "$skillRoot\examples\case-09-eight-pages.json" `
  -OutputVsdx (Join-Path $desktop ($runName + '-atlas.vsdx')) `
  -PreviewDirectory (Join-Path $desktop ($runName + '-previews')) `
  -Visible -KeepOpen
```

这会显示绘图过程，保存八页图册、导出每页 PNG，并保留成功生成的最终窗口。需要后台执行时去掉 `-Visible`；不需要保留窗口时去掉 `-KeepOpen`。

### 常用参数

| 参数 | 用途 |
| --- | --- |
| `-PlanJson` | 必填，JSON 绘图计划；支持一个文件、PowerShell 路径数组或含 `pages` 的图册计划 |
| `-OutputVsdx` | 必填，输出 VSDX 路径；不覆盖已有文件 |
| `-TemplateVsdx` | 可选，打开模板副本后绘制，保留原模板 |
| `-Visible` / `-KeepOpen` | 分别控制显示过程、成功后保留最终窗口 |
| `-PreviewDirectory` | 可选，导出所有页面 PNG |
| `-TimeoutSeconds` | 默认 300 秒；复杂图可按需延长，例如 `-TimeoutSeconds 600` |
| `-RunDirectory` | 可选，指定新的日志目录；默认位于输出旁的 `.visio-runs/<id>` |

直接运行脚本时，你需要自己准备 JSON 计划；它不会直接把 PNG 路径当成绘图计划。参考 [计划格式](skill/visio-template-drawing/references/drawing-plan-schema.md) 和 [公开示例](skill/visio-template-drawing/examples)。

<a name="faq"></a>

## 常见问题

**为什么没有出现 Visio 窗口？**

默认就是隐藏绘制。成功后打开输出 VSDX 查看；希望生成后自动显示，用 `-KeepOpen` 或在中文请求里明确说明。

**图可以继续编辑吗？**

主要内容由原生形状、文字和连接线组成，不是把整张参考图贴进 Visio。参考中的照片等栅格内容如需保留，仍可能以图片对象导入。

**输出文件默认保存在哪里？**

通过技能绘图时，本地参考文件默认保存到参考文件旁，文件名为 `<原名>_codex.vsdx`；仅粘贴在对话里的图片默认输出到桌面。你指定的位置和文件名优先。直接运行脚本时由必填参数 `-OutputVsdx` 决定。

**任务失败后如何排查？**

保留该任务的日志目录，查看 `result.json` 中的原始异常和清理结果。已经保存的最终文件不会因为后续导出失败而被删除；可能存在的 `failed-draft.vsdx` 是未完成现场，不代表任务成功。向维护者反馈前，检查日志中的私人路径和文档信息。

**窗口关闭了，为什么还有日志和测试文件？**

关闭任务所属窗口和进程，不等于删除磁盘文件。输出、日志和失败现场会保留供检查；确认不再需要后再按自己的备份策略清理。

**强制中断后能自动恢复吗？**

独立监护进程处理取消、超时及已确认归属的隐藏实例遗留。可见、明确保留或归属不明的实例不会被强行结束，以免影响用户文档。激活时尚未写入身份记录的极短窗口、整机断电和各种第三方插件卡死不能保证自动恢复。详见 [生命周期与恢复](skill/visio-template-drawing/references/session-lifecycle.md)。

**在哪里检查 COM 或已有文件的导出？**

使用 [visio_probe.ps1](skill/visio-template-drawing/scripts/visio_probe.ps1) 检查 Visio，使用 [visio_export_preview.ps1](skill/visio-template-drawing/scripts/visio_export_preview.ps1) 导出已有文档的指定页。两个入口都使用同一套实例归属和清理流程。

## 效果示例

| 参考图 | Codex 生成的可编辑 Visio 效果 |
| --- | --- |
| <img src="docs/assets/examples/example-01-reference.png" width="360" alt="Example 1 reference diagram"> | <img src="docs/assets/examples/example-01-codex.png" width="360" alt="Example 1 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-02-reference.png" width="300" alt="Example 2 reference diagram"> | <img src="docs/assets/examples/example-02-codex.png" width="300" alt="Example 2 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-03-reference.png" width="360" alt="Example 3 reference diagram"> | <img src="docs/assets/examples/example-03-codex.png" width="360" alt="Example 3 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-04-reference.png" width="330" alt="Example 4 reference diagram"> | <img src="docs/assets/examples/example-04-codex.png" width="330" alt="Example 4 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-05-reference.png" width="220" alt="Example 5 reference diagram"> | <img src="docs/assets/examples/example-05-codex.png" width="220" alt="Example 5 Codex-generated Visio output"> |

## v1.0.0：后台绘图与资源清理

- 单页绘图、按计划生成多页图册、导出和探测统一使用受管 Visio 实例。
- 显示过程与结束后保留窗口分开控制；默认隐藏并自动关闭。
- 增加结构化日志、原始异常记录、COM 释放、退出核验、取消、超时与严格限定的遗留恢复。
- 已实测单页和八页图册连续两轮、原生编辑与粘连、保存/导出失败、已有未保存文档保护、显示/保留模式、文件锁、取消和强制中断。用户并行操作其他软件的点击、输入、复制粘贴也已实际验证。
- 已验证环境：Visio 16、Windows PowerShell 5.1 STA；PowerShell 7.6 作为启动入口。未将所有机器和插件环境视为已验证。

完整更新说明见 [发行版](https://github.com/zxc-heu/visio-template-drawing-skill/releases/tag/v1.0.0)。

## 开发与验证

```powershell
python skill\visio-template-drawing\scripts\validate_plan_layout.py `
  skill\visio-template-drawing\examples\case-09-eight-pages.json

powershell -NoProfile -STA -File skill\visio-template-drawing\scripts\test_visio_lifecycle.ps1 `
  -OutputDirectory .\outputs\acceptance-new -IncludeVisible

powershell -NoProfile -STA -File skill\visio-template-drawing\scripts\test_visio_interruptions.ps1 `
  -OutputDirectory .\outputs\interruptions-new
```

每次测试使用新的输出目录。`-IncludeVisible` 会显示受控测试窗口；测试不拿用户真实未保存文档冒险。真实并行输入和复制粘贴仍需单独人工验证，不能由语法或布局检查代替。

可发布技能源码位于 `skill/visio-template-drawing/`。本地日志、一次性诊断、会话和生成的图册不纳入发布，发布前检查见 [PUBLISHING.md](PUBLISHING.md)。

## License

MIT，见 [LICENSE](LICENSE)。
