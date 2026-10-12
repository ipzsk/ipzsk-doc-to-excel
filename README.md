# 文档转 Excel · DSH 独立插件

原生 DSH / DeepSeek Harness Desktop 插件，提供「文档转 Excel」独立面板。无需 Skill 提示词或聊天操作。上传 .docx / .pdf，选择 Harness 已配置模型或纯本地规则，下载分类 Excel。

## 功能定位

- 小文档模式：批量读取 Word/PDF，按本地规则或模型要求分类，输出总览、分类明细、待确认和原始表格。
- 大 PDF 模式：支持最高 2 GB 的单个 PDF，先试跑 10 页，再按保存的列名和整理要求继续全文。
- 本地解析与 OCR：正文、表格和无文本页优先在本机处理；macOS 可调用 Vision OCR。
- 节省 token：去重、摘录预算、批处理、缓存、暂停与恢复，避免把全文、聊天历史和工具列表重复发送给模型。
- 数据留在当前 DSH Profile，输出 Excel 字段按文本保存，保留前导零和公式样式原文。
- 当前版本负责提取和生成 Excel，不执行数据库入库；正式入库前需要人工核对字段与原文。

## 与桌面端和插件市场的关系

本仓库独立于 [desktop-ipzsk](https://github.com/ipzsk/desktop-ipzsk) 发布。插件在本机完成文档解析；用户选择模型时，通过 DeepSeek Harness 的模型服务和 [token-ipzsk](https://github.com/ipzsk/token-ipzsk) 产生对应调用与计费记录。

新版本通过本仓库 GitHub Release 发布，再由开发者中心提交审核。上架后，桌面端的 IPZSK 插件市场读取 `/api/me/plugin-market`，下载审核附件并校验包名、版本、大小和 SHA-256。插件数据目录不随安装包替换，桌面端与插件可以分别更新。仓库创建或发布 Release 不会自动上架。

当前版本：[v0.2.1](https://github.com/ipzsk/plugin-dsh-doc-to-excel/releases/tag/v0.2.1)。

## 六个仓库如何协作

| 仓库 | 职责 |
| --- | --- |
| [web-ipzsk](https://github.com/ipzsk/web-ipzsk) | 官网、账户、开发者中心和 Web 创作工作台 |
| [desktop-ipzsk](https://github.com/ipzsk/desktop-ipzsk) | DeepSeek 官方底座、IPZSK 登录和插件市场宿主 |
| [token-ipzsk](https://github.com/ipzsk/token-ipzsk) | 账户、号池、模型路由、计费和插件审核 |
| [plugin-dsh-imagegen](https://github.com/ipzsk/plugin-dsh-imagegen) | AI 生图与改图插件 |
| [plugin-dsh-doc-to-excel](https://github.com/ipzsk/plugin-dsh-doc-to-excel) | 当前仓库，Word/PDF 转分类 Excel 插件 |
| [plugin-dsh-swf-studio](https://github.com/ipzsk/plugin-dsh-swf-studio) | FloorLED 与 SWF 场景制作插件 |

## 安装

需要支持 Host `webServer`、`llm` 与 Web Client 插件的 DSH 0.1.2 系列宿主。

```sh
dsh plugin --profile desktop add https://github.com/ipzsk/plugin-dsh-doc-to-excel/releases/download/v0.2.1/dsh-doc-to-excel-0.2.1.tgz --ignore-scripts --config.auto-install-peers=false
```

使用 Web profile 时把 desktop 改成 web。安装后重新加载 Harness。插件随 bundle patch 加载，侧边栏出现「文档转 Excel」。仅支持本机同源访问。

## 本地处理环境

需要 Python 3.10+，以及 python-docx、pypdf、openpyxl。插件优先使用 `DSH_DOC_EXCEL_PYTHON` 指定的解释器，然后检查插件 `.venv`、本机 Codex 已有文档运行环境与 PATH 中的 Python；不会自动下载依赖。

如果本机没有这些库，可自行创建文档专用环境：

```sh
python3 -m venv ~/.dsh/doc-excel-python
~/.dsh/doc-excel-python/bin/python -m pip install python-docx pypdf openpyxl
```

启动宿主时设置 `DSH_DOC_EXCEL_PYTHON` 为该环境解释器的绝对路径。Windows 对应 `Scripts/python.exe`。密钥由 Harness 模型提供方管理，本插件没有独立密钥配置，也不会读取密钥文件。

## 小文档模式使用

1. 侧边栏打开「文档转 Excel」，上传 1–10 个文档，合计最多 20 MB。
2. 选择模型。默认纯本地规则，不向模型发送内容。选模型后，仅把规则未处理的摘录发送到该模型服务。
3. 可自定义分类，每行 `分类名称：关键词1,关键词2`；留空使用默认分类。
4. 设置摘录预算并开始整理。结果包含总览、分类明细、待确认、原始表格四个工作表。

## 节省 token 的方式

- Python 本地解析和导出，不把全文、聊天历史、工具列表交给模型。
- 关键词唯一命中直接分类；重复内容仅分类一次，每处来源都保留。
- 每批最多 12 条，每条摘录 360 字符，每批最多 6000 字符；单任务摘录预算默认 12000 字符。
- 最多 12 次分类调用，每次分类输出上限 3200 token；模型支持时选择最低档推理。
- 插件不自动重试失败调用；宿主适配器自身的网络重试策略仍由 Harness 管理。
- 同文件内容、分类配置、预算和模型的成功结果会复用磁盘缓存。缓存命中本次不再调用模型，界面显示原任务累计统计。
- 字符统计不是账单 token。模型返回 usage 时显示实际报告用量，缓存 token 分开统计；异常中断可能没有完整 usage。
- 分类只依据摘录，信息不足时降低置信度。预算不足、模型失败或用户停止后，保留结果供人工确认。

## 数据与边界

文档和结果留在当前桌面 profile 的 `doc-to-excel-jobs`；无 Desktop profile 服务时放在 `~/.dsh/doc-to-excel-jobs`。不会自动清除原始文件。重新上传相同输入可找回缓存结果。

Word 提取正文段落与表格，不含页眉页脚、文本框、批注。macOS 且安装 Swift 命令行工具时，使用系统 Vision 在本地识别 Word 正文内嵌图片和无文本 PDF 页；每文档最多 30 张图片或 30 个无文本页，不消耗模型 token。首次使用会在本机编译并缓存 OCR 程序。其他系统或识别失败时保留待提取标记。OCR 内容始终进入待确认。PDF 按页提取文本；有文本的页面不额外识别嵌入图片。复杂 PDF 的多栏顺序和表格结构需复核。预览显示前 100 条、每条 300 字符；Excel 保留全部已提取文本。公式样式的输入按普通文本写入，避免意外执行公式。

当前面板使用与现有 imagegen 插件一致的侧栏 DOM 挂载方式，宿主更改侧栏结构后可能需要适配。插件不更改桌面应用二进制或上游源码。

## 验证

```sh
node --test tests/engine.test.js
/path/to/python -m unittest discover -s tests -v
```

Node 测试覆盖本地模式、去重、模型校验失败停止、取消和磁盘缓存；Python 测试覆盖真实 Word/PDF、表格、公式文本、长内容、预算和低置信度。测试源码位于完整项目目录；分发包仅包含运行代码。


## 0.1.3：像聊天一样描述整理要求

面板新增「整理要求」：写出提取目标、分类方式、保留/排除条件。填写后必须选择模型；可在「指定分类」每行指定一个类目，否则额外用一次小请求根据要求设计分类。此模式跳过关键词自动判定，让模型按用户要求处理所有唯一文本块。长内容按最多 1000 字符分段送入预算范围内的批次，避免仅凭整页开头概括整页。

模型输出整理内容、分类理由、置信度和保留标记。分类明细同时保存整理内容和原始提取文本；高置信度排除项放入「已排除」，低置信度及 OCR 项保留待确认。预览可展开原文和分类理由。修改要求后点击「按当前要求重新整理」，相同文件不必再次选择；要求纳入缓存标识。

小文档模式仍采用固定工作簿；0.2.0 大 PDF 模式支持填写列名。OCR 识别错误仍需原图核对，不让模型猜补数字或名称。每任务最多 12 次调用，包含可能的分类方案请求；分类输出上限现为 3200 token，方案上限 2200 token。比纯规则模式用量增加，但没有聊天历史与工具循环。


## 0.2.0：大 PDF 按列提取

默认进入大 PDF 模式，小文档原有功能保留在模式选择器中。

1. 选择单个 PDF（上传上限 2 GB），选择 Harness 已有模型。若列表存在 Qwen3.8，会优先选中；请核对具体版本和渠道。
2. 填写整理要求、每行一个 Excel 列名（最多20列）。留空输出整理内容。
3. 点击「上传并试跑10页」：4 MB 二进制分块上传到本机，随后解析并调用所选模型。浏览器不会整份 Base64 编码文件。
4. 核对预览，可导出样本。点击「继续全文」沿用保存的模型、列名和要求。变更这些设置需要创建新任务。
5. 本轮字符额度耗尽或模型错误时暂停；继续会授予新的额度。退出应用后，通过已保存任务选择器恢复。上传中断的任务当前需要重新上传，已完成上传的处理任务支持恢复。
6. 下载 Excel：提取结果、待核对、原文、页面异常、总览。字段强制为文本，保留前导零和公式样式原文，不执行数据库入库。

实现：Python 标准库 SQLite 逐页事务提交，单次读取 PDF 页面树；模型每次最多两个1800字符片段，前后各250字符仅辅助理解，不发送聊天历史。没有 OCR 30页限制，macOS 无文本页逐页尝试本地 Vision。结果原始 JSON 先保存再应用；已成功结果不会因正常恢复重新请求。输出每批上限8000 token；用户明确继续才重试模型失败项。宿主网络重试可能独立存在。字符额度非费用硬上限，不包含输出token；异常调用可能已计费但没有 usage。每轮额外限制最多1000次模型调用。

完整性：未解析页、解析失败页、未整理片段均有计数。所有页面成功解析且片段处理完毕后才显示全文处理完成；这表示处理覆盖完成，不代表字段内容已人工验证。OCR、模型核对说明、没有匹配记录和未处理项保留在输出。跨页上下文不等同于可靠表格重建/记录合并；复杂表格需样本验证。带文字页面内的图片不会额外OCR。本版没有自动入库、自动唯一键去重、Excel模板类型校验或视觉模型复核。

任务存放在 profile 的 `doc-to-excel-jobs/large/<id>`，包含源文件、SQLite、固定配置、模型响应及 Excel。未自动清理，占用磁盘随源文件和文本量增长。处理中可关闭面板，电脑休眠会影响进度；退出宿主会中止，重启后手动继续。并发为一个大 PDF 任务，防止同时 OCR 和模型请求拖垮本机。2 GB 是上传限制，尚未验证所有接近2GB的PDF；特别复杂/损坏PDF仍可能超过本地处理限制。

验证：`node --test tests/*.test.js`；Python 文档测试同上；显式规模测试 `/path/to/python tests/large_stress.py` 生成约30 MB、4000页的合成文字PDF，核对逐页恢复和Excel原文行数。模型端到端逻辑通过可控模型响应验证，真实业务质量仍需实际PDF样本确认。


## GitHub Release 与插件市场更新（0.2.1 起）

仓库：https://github.com/ipzsk/plugin-dsh-doc-to-excel

发布：https://github.com/ipzsk/plugin-dsh-doc-to-excel/releases

IPZSK Desktop 用户通过插件市场检查并安装后台审核版本。其他兼容 Harness 可以在插件面板顶部点击「检查插件更新」；更新器只读取本仓库的最新正式 Release，只接受固定命名的安装包与 SHA256SUMS，校验成功后调用当前 profile 的 `dsh plugin add`，禁止安装脚本，安装完成核对版本并提示重启。更新过程不会发送文档、模型密钥或聊天内容到 GitHub，也不会更新桌面应用本体。

自动安装要求客户端 Host 提供当前 profile 目录，且 `dsh` 与 `pnpm` 命令可用。没有该运行环境时，先从 Releases 下载安装包，在客户端终端执行同上安装命令（替换成目标版本链接）。GitHub 网络失败、校验失败会停止；源码开发链接会被明确识别为未成功切换，不能冒充升级成功。

本地源码开发安装切换为发布包：先用目标版本的安装命令更新依赖声明；如果 pnpm 提示仍链接源码，再运行 `dsh plugin --profile desktop unlink dsh-doc-to-excel`，然后重启。其他 profile 替换 desktop。该命令解除包链接，不删除源码目录或任务数据。

更新前暂停文档处理。已有任务数据位于 profile 的独立目录，包更新不删除它们。安装失败时可重新安装此前的 Release 包；自动回滚尚未实现。跨版本的数据结构兼容性会在相应版本发布说明中注明；请不要在未来引入数据迁移后盲目降级。

### 维护者发布

修改 package.json 的版本号，执行测试并提交源码；推送相同的版本标签（例如 v0.2.2）。GitHub Actions 会先执行 Node/Python 测试，再创建 Release，上传 tgz 和 SHA256SUMS。创建 Release 后还需在 IPZSK 开发者中心提交审核；审核通过才会出现在插件市场。通用 Harness 客户端只在用户点击检查时请求 GitHub，不后台轮询、不自动升级。

也可用 `pnpm release:pack` 在本地生成 `release/` 下的相同两个文件。构建仅使用 Node 和 pnpm，不需要复制本机配置。仓库和安装包不包含业务样本、Excel 输出、任务数据库、模型账号配置或 .env。
