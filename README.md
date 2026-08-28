# ExcelAI

ExcelAI 将多模型 Agent 能力接入 Excel，当前内置 DeepSeek V4、SiliconFlow、OpenRouter 与自定义兼容网关。支持 Microsoft 365 管理员集中部署，以及使用独立加载项 ID 的 Windows EXE 本机兼容安装。

## 功能

| 功能 | 说明 |
|---|---|
| Excel 上下文 | 当前单元格/选区（默认）、当前工作表、整个工作簿三种模式，聊天工具栏切换，发送时自动附加 |
| 模型与推理等级 | deepseek-v4-flash / deepseek-v4-pro；思考模式开关 + 推理强度（低/高/最大），均在聊天工具栏 |
| API 接入与管理 | 右上角 ⚙ 设置弹窗：官方 / SiliconFlow / OpenRouter / 可命名的多个自定义供应商；自定义名称、base_url、模型均持久化，可继续添加与切换；API Key 保存/清除；连接测试 |
| Agent 读写与格式 | 工具调用循环；读类工具（选区/工作表/指定区域 read_range/透视表/图表/名称等）自由调用；写值/公式与格式设置（填充色/字体色/加粗/隔行相间着色）均需页面内确认弹窗 |
| 数据透视表 | list_pivots / create_pivot（两步创建；字段通过官方推荐的 hierarchies 集合获取与添加——hierarchies.getItem 传给 rowHierarchies/dataHierarchies.add；读取不到时降级读源区域表头；逐字段容错；创建失败时列出已有透视表并提示换位置）/ read_pivot / refresh_pivot（单个或全部） |
| 数据连接 | refresh_connections：一键刷新全部数据连接（含 Power Query 查询结果）；Power Query 查询本身的创建/编辑 Office.js 不支持 |
| 菜单功能扩展 | 条件格式（cellValue 规则/colorScale 色阶/dataBar 数据条/iconSet 图标集）、图表（创建含图例开关/删除/列出）、表格（创建/列出，名称自动清洗）、排序、自动筛选、冻结窗格、数据验证（下拉/整数/小数/日期）、批注（增删）、迷你图、名称管理器（定义/列出/删除）、行高列宽（自动调整）、合并/拆分单元格、清除内容/格式、隐藏行/列、工作表管理（新建/重命名/删除）、打印设置（打印区域/方向/缩放/按页适配）；创建/删除/修改类弹确认，排序/筛选/冻结/行高列宽/隐藏/打印为可逆操作不弹框 |
| 对话记录 | 侧边栏多会话（新建/切换/双击重命名/删除）+ 🔍 搜索（按标题或消息内容实时过滤），localStorage 持久化；设置弹窗内导出 JSON（复制）/导入/导出到工作表 |
| 外观 | 明/暗主题（右上角切换）、主色（蓝/紫/绿/橙）、字号（小/中/大），在设置弹窗 |
| 技能 | 内置技能（数据分析师/公式助手/VBA 助手）+ 自定义技能；聊天工具栏下拉选择；设置弹窗内添加/删除/导出导入（JSON） |
| 文件上传 | 聊天工具栏 📎 附件按钮：txt / csv / md / json / xlsx（SheetJS 解析），图片暂不支持 |
| 其他 | 每条消息右下角带 📋 复制按钮；**工具调用日志默认折叠**（摘要行只显示 🔧 工具名，点击展开查看完整参数与结果）；模型输出支持 Markdown H1–H6 标题、列表、引用、分隔线、表格、粗体、斜体、行内代码和代码块；**快捷操作显示在新建对话的空白区**（选择后消失，设置弹窗「⚡ 快捷指令」可自定义增删）；**权限开关在聊天工具栏**（🔐 请求批准 / ⚡ 替我批准）；Ctrl/Cmd+Enter 强制发送，中文输入法组合中回车不误发送；会话列表 ✎ 重命名；阅读历史时新消息不强制拽到底部；发送时日志区显示"已附加：上下文 X KB · N 个文件"；存储空间不足时自动提示导出备份（设置内显示占用）；⏹ 停止生成；侧边栏默认收起；上下文智能截断（头尾保留、选区 300×100 上限）；🧰 工具清单；完整版本记录见 CHANGELOG.md |

**功能区按钮**：`manifest.xml` 用于 Microsoft 365 管理员集中部署；`manifest-standalone.xml` 用于 Windows EXE。两者使用不同加载项 ID，可同时保留且不会互相覆盖。本机 EXE 通过应用私有隐藏载体尝试在 Excel 启动时激活本机清单；桌面快捷方式可为已经运行的 Excel 会话补载。该方式属于兼容方案，不等价于 Microsoft 365 的正式用户授权；Office 组织策略、受保护视图或安全模式仍可能阻止本机侧载。

**界面结构**：主界面为纯聊天框；聊天框上方任务栏提供模型、推理强度、思考模式、上下文范围、技能、附件、自动附加开关；右上角 ⚙ 设置采用分类导航，收纳 API 接入、技能管理、快捷指令、外观与数据管理；对话记录侧边栏默认收起。

## 免本地服务器部署（推荐：HTTPS 静态托管）

**彻底去掉本地服务器**：把页面托管到公网 HTTPS 静态空间后，无黑窗、无空闲退出，且支持 Excel 网页版/Mac（API Key 仍只存在你本机的 Office WebView 配置中，静态空间不涉及任何密钥）。GitHub Pages 只负责托管页面、JavaScript、CSS 和图标，不负责将加载项授权给某个 Office 用户。

**GitHub Pages 完整步骤（免费）：**

1. 注册/登录 GitHub → 新建**公开**仓库，名字如 `deepseek-excel-assistant`（勾选 Add a README）
2. 仓库页面 → **Add file → Upload files** → 把 `pages\` 文件夹里的 **8 个文件**拖进去 → Commit
3. 仓库 **Settings → Pages** → Source 选 **Deploy from a branch** → Branch 选 **main /(root)** → Save，等约 1 分钟得到地址 `https://你的用户名.github.io/deepseek-excel-assistant/`
4. 双击 **`部署到GitHubPages.bat`** → 粘贴你的 Pages 地址 → 自动切换清单并同步到已安装目录（也可手动：`node deploy.mjs 地址`）
5. 完全退出 Excel → 打开普通工作簿 → 点「开始」选项卡的 ExcelAI 按钮 → 确认 v0.39 页面正常打开
6. 公网 HTTPS 模式不需要 Node.js 或本地服务。

> 更新版本时：把新的 8 个文件重新上传到仓库覆盖即可（页面 URL 不变）。

### 两种分发方式可以同时保留

- **组织内用户（推荐）**：管理员进入 **Microsoft 365 管理中心 → 设置 → 集成应用 → 上传自定义应用**，上传 `manifest.xml`，并分配给用户或组。目标用户登录受支持的 Office 后，由 Microsoft 365 下发授权和清单。
- **其他 Windows 用户**：运行 `dist\ExcelAI-Standalone-Setup-v0.39.exe`。安装器使用 `manifest-standalone.xml` 和独立本机 ID，不会覆盖集中部署版。

GitHub Pages 只托管任务窗格页面、脚本、样式和图标。无网络或无法访问该站点时，按钮可能仍显示，但任务窗格内容无法加载。它不承担加载项授权和分发。

## EXE 本机兼容安装

1. 关闭所有 Excel 窗口并运行安装包；安装为当前用户，不需要管理员权限。
2. 重新打开 Excel 并新建或打开普通工作簿。若当前 Office 允许本机 Developer 侧载，开始选项卡会出现 **ExcelAI（本机）**。
3. 若按钮未出现，可运行桌面快捷方式补载当前会话。仍不出现通常表示组织策略限制侧载；此时应改用 Microsoft 365 管理员集中部署，EXE 无法绕过该策略。

安装器不会创建网络共享，也不要求打开 `deepseek-excel-assistant-sideload.xlsx`，默认 HTTPS 模式不启动 Node.js 或其他后台服务。

## 安装包（exe，基于 Inno Setup）

**已生成**：`dist\ExcelAI-Standalone-Setup-v0.39.exe`

**其他机器安装步骤**：
1. 关闭 Excel，双击安装包；安装器为当前用户复制文件、注册清单、安装隐藏启动载体并创建新图标快捷方式。公网 HTTPS 模式下快捷方式不启动 Node.js 或其他后台服务；如 Excel 已在运行，它会显式补载隐藏载体。
2. 完全重启 Excel并打开普通工作簿；页面加载自 GitHub Pages，本地无 Web 服务。功能区是否允许本机侧载仍受 Office 版本与组织策略控制。

- 卸载时自动清理：本加载项的注册、功能区缓存、目录与快捷方式
- 安装路径可选择（默认仍为 `%LOCALAPPDATA%\DeepSeekExcelAssistant`，用于兼容旧版原位升级）
- 重新生成：双击 `生成安装包.bat`（自动调用 Inno Setup 编译器 ISCC.exe），或手动运行 `ISCC.exe setup.iss`

## 快速开始

**前提**：Microsoft 365 桌面版 Excel。只有清单指向 localhost 时才需要 Node.js；默认 HTTPS 模式不需要。

1. 二选一：管理员集中部署 `manifest.xml`，或运行 Windows 独立安装包；开发目录使用时可运行 `安装注册.bat`。
2. 完全重启 Excel，打开任意普通工作簿。
3. 点击「开始」选项卡中的 ExcelAI 按钮。
4. 任务窗格中打开「⚙ 设置」→ 填 API Key → 测试连接 → 开始对话。

## 安装机制说明

- 集中部署清单 ID：`7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c`；EXE 本机清单 ID：`1537f254-10aa-41d5-aed2-0a00b89da104`。安装器升级旧版本时仅删除旧的 `WEF\Developer` 映射，不会删除 Microsoft 365 管理中心的部署关系。
- 安装器注册 `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\1537f254-10aa-41d5-aed2-0a00b89da104`，指向安装目录中的 `manifest-standalone.xml`。
- `DeepSeekExcelAssistant-Standalone-Autoload.xlsm` 安装到应用私有的 `carrier` 子目录；安装器只将该专用目录登记为 Excel 受信任位置，并分配一个不占用现有配置的 `OPEN/OPENn` 启动项。启动参数使用 `/r` 强制只读，因此多个独立 Excel 进程可以同时载入。
- 工作簿窗口为隐藏状态，只负责在 Excel 启动时激活 WebExtension；不会占用或改写用户已有的备用启动目录。
- 载体内只有一个 `Workbook_BeforeClose` 清洁事件，用于清除 Excel 因加载嵌入式任务窗格产生的“伪修改”标记，避免退出时反复询问保存；它不会访问或改写用户工作簿。
- `deepseek-excel-assistant-sideload.xlsx` 只保留作开发测试文件，不是用户启动入口。

## 技术要点（已踩坑记录）

- DeepSeek V4：`thinking: {"type":"enabled"}` + `reasoning_effort`（low/high/max）；带工具调用必须回传 `reasoning_content`；`deepseek-chat/reasoner` 旧名已停用。
- CORS：三个供应商均放行浏览器直连，无需本地代理。
- Office 任务窗格禁用 `window.confirm`，写入确认用页面内弹窗实现。
- 思考链流式片段极小，需聚合显示（否则逐字换行）。
- 发送/测试前从输入框同步最新 Key/自定义地址/模型（避免改了没点保存用旧值）。

## 自动化测试

```
node excel-addin/test/app.test.mjs
```

用 vm 沙箱加载真实 `app.js`（模拟 DOM/localStorage/fetch/Excel），163 项断言覆盖：
md 渲染（公式不竖排回归）、SSE 流式装配（思考链+工具调用片段+finish_reason）、Agent 完整循环（工具结果与 reasoning_content 入历史）、写值确认弹窗（允许/拒绝）、格式设置（填充/字体/加粗/隔行相间偶数行着色/拒绝路径）、停止生成（中断 + 友好提示）、上下文截断策略（trimContext 头尾保留 / capGrid 行列上限）、文件附加（含 xlsx 无 SheetJS 优雅降级）、导出到工作表、localStorage 持久化、完整发送流程（上下文附加/文件注入/标题更新）、API 错误路径（401/402/400/网络失败友好提示）、技能（默认/注入/自定义添加/删除/导入导出）、主题切换、导入对话、历史渲染、会话切换与删除、侧边栏折叠、设置弹窗内结果展示、数据透视表（两步创建/字段匹配容错/双重失败友好错误/表头提示/读取/刷新/列出）、数据连接刷新、菜单功能扩展全链路、read_range 通用区域读取、表格名清洗与参数守卫、工具 Schema 合法性与名称唯一性、复制按钮、权限设置（自动批准/请求批准）、工具清单弹窗。

## Git 工作流

本目录已初始化为 Git 仓库并连接 `origin`：`https://github.com/bjzzxhzl/deepseek-excel-assistant.git`。`main` 根目录同时作为 GitHub Pages 发布源；`dist`、`outputs`、`pages` 与临时测试文件由 `.gitignore` 排除。开发完成后运行测试、生成安装包，再提交并推送 `main` 即可同步源码与 Pages。
