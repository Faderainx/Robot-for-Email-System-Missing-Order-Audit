# 邮件&系统漏单审核自动化工具 — 项目文档

## 项目位置

- **项目根目录**: `b:\TRAE_Project\6aa0bf7bc4ecce8d9c80e568\mail_audit_bot\`
- **设计文档**: `b:\TRAE_Project\6aa0bf7bc4ecce8d9c80e568\design_plan.html`
- **Git 仓库**: 项目根目录下 `.git/`，分支 `master`
- **Python 环境**: `C:\Users\35144\AppData\Local\Programs\Python\Python312\python.exe`
- **启动文件**: `b:\TRAE_Project\6aa0bf7bc4ecce8d9c80e568\mail_audit_bot\启动.bat`

## 业务目标

将"邮件&系统漏单审核"流程自动化。原流程完全依赖人工：逐封阅读邮件，识别注册/新增/撤单类邮件，提取关键字段（代理、客户、项目），再在工单系统中逐条检索是否已录单。未录单的标记为漏单。工具自动完成邮件读取→过滤→字段提取→项目标准化→工单比对→输出 Excel 全流程。

## 文件结构

```
mail_audit_bot/
├── main.py                     # 程序入口
├── gui.py                      # PyQt5 GUI 界面
├── config.yaml                 # 配置文件 (邮箱/工单/LLM 凭据)
├── requirements.txt            # Python 依赖清单
├── 启动.bat                    # Windows 启动脚本
├── .gitignore                  # Git 忽略规则
├── test_modules.py             # 逐模块测试脚本
│
├── modules/                    # 核心业务模块
│   ├── mail_reader.py          # M1 邮件读取 (IMAP 阿里邮箱)
│   ├── mail_filter.py          # M2 邮件过滤 (规则引擎 + LLM 二次识别)
│   ├── llm_intent.py           # LLM 调用、一次修复与调用审计
│   ├── agent_schemas.py        # Pydantic 严格输出协议
│   ├── agent_workflow.py       # 本地 Agent Workflow 编排
│   ├── business_validator.py   # 明显异常字段的确定性业务校验
│   ├── field_extractor.py      # M3 字段提取 (表驱动 + 模糊匹配)
│   ├── project_normalizer.py   # M4 项目标准化与拆分
│   ├── workorder_checker.py    # M5 工单比对 (Playwright 半自动)
│   └── excel_writer.py         # M6 双 Excel 输出 (带颜色标记)
│
├── utils/                      # 工具模块
│   ├── fuzzy_match.py          # 模糊匹配 (OpenCC 简繁 + rapidfuzz)
│   ├── attachment_parser.py    # 附件解析 (zip/rar/xlsx/xls/pdf/docx/图片OCR)
│   └── logger.py               # 日志工具
│
├── data/                       # 参照表 Excel (人工维护)
│   ├── internal_emails.xlsx    # 附件二: 公司内部邮箱表
│   ├── agent_emails.xlsx       # 附件三: 代理邮箱对照表
│   └── project_names.xlsx      # 附件四: 标准项目名称表
│
├── output/                     # 输出结果目录
│   ├── stage1_email/                     # 阶段一：邮件解析结果
│   │   ├── to_workorder_list.xlsx        # 可进入阶段二工单查询
│   │   ├── to_review_list.xlsx           # 需人工复核/补全
│   │   ├── filtered_mail_record.xlsx     # 过滤邮件审计
│   │   └── *_YYYYMMDD_HHMMSS.xlsx        # 阶段一历史副本
│   ├── stage2_workorder/                 # 阶段二：工单系统比对结果
│   │   └── workorder_check_result*.xlsx
│   ├── manual_review/                    # 工作台人工确认后的结果
│   └── diagnostics/                      # 预留：探测文件、调试截图等
│
├── cache/                      # 本地缓存目录 (可清理, 不影响业务)
│   ├── mails/{mailbox}_{uid}.eml          # 邮件原始字节 (按邮箱+UID 命名)
│   └── attachments/{hash}_{filename}      # 图片附件按内容 hash 去重
│
├── storage/                    # 持久化数据目录
│   ├── query_cache.json        # 工单查询缓存 (代理|公司|项目 三联键, TTL 默认 7 天)
│   └── workbench.db            # 独立工作台 SQLite 数据库（输入增量 + 操作留痕）
│
├── logs/                       # 运行日志目录
└── session_state.json          # GUI 状态持久化 (时间范围/汇总表路径/运行模式/阶段二输入路径)

```

## 缓存与持久化

### 独立工作台数据库

工作台可以通过 `启动独立工作台.bat` 或 `python workbench_launcher.py` 单独启动，不需要打开 PyQt 主程序。阶段一/阶段二每次产生的新数据会增量写入 `storage/workbench.db`；工作台启动或点击“刷新数据”时优先读取数据库，Excel 仍作为兼容输入并会自动导入。数据库按邮件身份幂等更新，历史数据不会因下一次运行被覆盖。

外部程序也可以向本机服务发送 `POST /api/ingest`，请求体使用 `{"dataset":"active","rows":[...]}`；可用 `GET /api/operations?date=YYYY-MM-DD` 按操作日期读取人工/系统留痕，`GET /api/database` 查看数据库统计。

| 类型 | 路径 | 作用 | 清理影响 |
|------|------|------|----------|
| 邮件字节缓存 | `cache/mails/{mailbox}_{uid}.eml` | IMAP 拉到原始字节即落盘，重跑同日期范围直接读本地，绕过阿里邮箱 IMAP 5 分钟超时 | 删除后下次同日期范围会重新拉取（IMAP `email.reconnect_batch_size` 阈值内分批） |
| 图片附件缓存 | `cache/attachments/{sha1[:12]}_{filename}` | 按内容 hash 去重，zip/rar 内图片同样落盘，OCR 兜底阶段才读取 | 删除后下次解析同一封邮件的同一图片会重新落盘（内容相同则 hash 相同，自动跳过） |
| 工单查询缓存 | `storage/query_cache.json` | 阶段二按 `代理|公司|项目` 三联键缓存 RPA 查询结果（命中条件/工单日期等），TTL 默认 7 天（`workorder.query_cache_ttl_days`） | 删除后下次阶段二会重新走 RPA 查全部条目 |
| GUI 状态 | `session_state.json` | 时间范围、邮箱汇总表路径、运行模式（阶段一/二/一站式）、阶段二输入文件路径 | 删除后下次启动恢复默认值 |

## 核心功能 — 六大模块

### M1 邮件读取模块 (mail_reader.py)

- **协议**: IMAP SSL，服务器 `imap.mxhichina.com:993`
- **邮箱**: 阿里企业邮箱 `huiyan.song@eu-helper.com`，使用客户端授权码登录
- **只读模式**: `conn.select("INBOX", readonly=True)`，不改变邮箱任何状态
- **分批拉取**: 每 20 封重连一次，避免阿里邮箱 IMAP 5 分钟超时踢人
- **日期兼容**: `fetch_mails(date_from, date_to)` 支持 datetime 对象和 `YYYY-MM-DD` 字符串
- **离线解析**: 先批量拉取邮件原始字节，断开 IMAP 后再逐封解析正文和附件
- **自身发送排除**: 发件人是 `huiyan.song@eu-helper.com` 的邮件标记 `skip_reason: "self_sent"`
- **附件解析支持**: .xlsx .xls .pdf .docx .csv .txt .zip .rar .jpg .jpeg .png .bmp .tiff

### M2 邮件过滤模块 (mail_filter.py)

**三层筛选机制**:

**第一层 — 规则引擎 (mail_filter.py)**

按顺序判断，命中即停止:

1. **排除自身发送**: `skip_reason == "self_sent"` → 过滤
2. **无关词排除**: 含"账单/合同/咨询/下证/保证金/担保/退款/回收费/购买"且不含业务词 → 过滤
3. **方法一**: 主题/正文含"注册/新增/撤单" + 含业务词(WEEE/电池法/包装法等) → 有效
4. **方法一扩展**: 主题含注册关键字 → 有效
5. **方法二**: 主题/正文/附件含业务关键字(WEEE/电池法/包装法/EPR/国家名等) → 有效
6. **内部无关**: 发件人在内部邮箱表且不含业务词 → 过滤
7. **外部无关**: 不含业务词也不含注册词 → 过滤
8. **不确定**: 以上都没命中 → 保留，标灰色

**关键字定义**:

| 类别 | 关键字 |
|------|--------|
| 动作词 (KEYWORDS_ACTION) | 注册、新增、撤单、注销、追加 |
| 业务词 (KEYWORDS_BUSINESS) | WEEE、电池法、包装法、EPR、一次性塑料、BAT、德国、法国、意大利、西班牙、荷兰、波兰、瑞典、比利时、爱尔兰、葡萄牙、奥地利 |
| 无关词 (KEYWORDS_IRRELEVANT) | 账单、invoice、合同、协议、咨询、反馈、建议、下证、证书号、保证金、担保、退款、回收费、购买 |

**第二层 — 意图 Agent (可选)**

- 对第一层标记为“不确定”的邮件，以及允许复检的过滤邮件，逐封做意图分类
- 规则明确命中的注册邮件不重复调用；硬过滤项不调用，控制时间和 Token 成本
- LLM 判定"注册类" → 恢复到有效列表 (标记"LLM恢复")
- LLM 判定其他 → 确认过滤
- API、JSON 或字段协议失败 → 保留原邮件并强制进入人工复核，不能视为“非目标”
- `api_key` 为空时自动跳过，仅用规则引擎

**第三层 — 人工复核**

- 输出两份 Excel: 漏单清单 + 过滤清单
- 颜色标记: 红色=漏单、黄色=日期异常、蓝色=待确认、橙色=多匹配、灰色=不确定

### 本地 Agent Workflow 与输出闸门

自动链路固定为：规则预筛 → 意图 Agent → Pydantic 校验 → 规则/字段 Agent 提取 →
Pydantic 校验 → 确定性业务校验 → 语义复检 Agent → Pydantic 校验 → Excel/人工复核。

- 不依赖 Dify 或 Coze，继续由现有 Python GUI 启动。
- 每个模型输出只允许一次格式/结构修复；再次失败就转人工。
- Pydantic 使用严格类型并禁止额外字段，避免字符串 `"false"` 被当成真值、项目字符串被当成数组等问题。
- 批量复检要求返回 ID 与请求一一对应，缺失、重复、越界都会失败并转人工。
- “一家公司”、说明句、邮箱、附件名等明显不是真实公司的内容由确定性规则拦截；系统只标记，不自动编造或改写公司名。
- EPR 申请表中的“POA法人职务/Legal positions”“注册资本/Registration Capital”“签字地点/时间”、
  联系人、证件号、地址、邮箱及表头说明不是客户明细；结构化解析会拒绝这些字段，字段 Agent
  无法确认公司主体时保持空值并转人工。旧数据库中的同类脏记录仅保留历史，不再显示在业务队列。
- 缺代理、客户、项目、需求，低置信，或任一校验未通过的行只能进入 `to_review_list.xlsx`，不能进入阶段二。

### M3 字段提取模块 (field_extractor.py)

**表驱动匹配，不依赖 LLM**:

| 字段 | 提取方式 | 依赖 |
|------|----------|------|
| 代理 | 发件人邮箱 → 附件三查表 (精确匹配 → 模糊匹配 → 多匹配全部展示) | 附件三 |
| 客户 (公司名) | 主题分隔符切分 → 正文正则 → 附件文件名 → 附件内容 | 邮件格式规律 |
| 项目 | 附件四标准项目名列表逐项扫描主题/正文/附件文本 | 附件四 |
| 需求 | 正则匹配"注册""新增""撤单" | 无 |

**代理提取流程**:

1. 精确匹配: 发件人邮箱在附件三中 → 直接取代理名
2. 正文邮箱: 从正文提取邮箱地址 → 查附件三
3. 模糊匹配: rapidfuzz 相似度 ≥ 80 → 候选
4. 多匹配: 相似度 ≥ 90 且多条 → 全部展示，标橙色交人工排查
5. 无匹配: 留空，标蓝色

**agent_map 数据结构兼容**:

- 字典格式: `{email: {"代理": str, "代理简称": str, "收件人邮箱": str}}`
- 字符串格式: `{email: "代理名"}`

### M4 项目标准化模块 (project_normalizer.py)

**组合项目拆分规则**:

- 分隔符拆分: 按 `+ ＋ / ／ 、 ， ,` 拆分
- 国家组合: "波兰荷兰包装法" → ["波兰包装法", "荷兰包装法"]
- N国模式: "荷兰、瑞典、波兰、比利时4国包装法" → 4 条
- 组合映射: "荷兰WEEE+包装法" → ["荷兰WEEE", "荷兰包装法"]

**SOP 示例验证**:

| 输入 | 输出 |
|------|------|
| 德国weee/电池/包装 | 德国WEEE + 德国电池法 + 德国包装法 (3条) |
| 波兰荷兰包装法 | 波兰包装法 + 荷兰包装法 (2条) |
| 荷兰、瑞典、波兰、比利时4国包装法 | 4 条独立项目 |
| 德国WEEE+电池法 | 德国WEEE + 德国电池法 (2条) |

**数据结构兼容**: 支持中文 key (`项目名称`) 和英文 key (`project_name`)

### M5 工单系统比对模块 (workorder_checker.py)

**半自动登录 (验证码人工输入)**:

1. Playwright 打开 Chrome 浏览器，导航到 `https://mp.ecopv-epr.com/workbench/main`
2. 自动填写账号 `Tommy` 和密码 `ba84ENj&H2`
3. 暂停等待用户输入验证码 (最长 120 秒)
4. 检测 URL 跳转离开 `/login` → 认为登录成功

**工单查询流程 (每条数据)**:

1. 清空上次查询条件 (输入框清空、下拉框重置)
2. 填入公司名称 (文本输入框)
3. 选择所属代理 (下拉框，支持原生 select 和自定义 el-select/ant-select)
4. 选择国家 (从项目名称提取，如"德国WEEE"→国家="德国")
5. 选择服务项目 (从项目名称提取，如"德国WEEE"→服务项目="WEEE")
6. 选择状态 (下拉框)
7. 点击查询按钮
8. 读取表格结果
9. 无结果 → 模糊匹配公司名称 (去掉"有限公司"等后缀取核心词重查，其他下拉不动)
10. 下一条数据 → 回到第 1 步

**国家提取**: 从项目名中匹配 24 个欧洲国家名

**服务项目提取**: WEEE / 电池法 / 包装法 / EPR / 一次性塑料

**四维模糊比对**: 日期(±60天) + 代理(相似度≥80) + 客户(相似度≥85) + 项目(相似度≥85)

**只读操作**: 只做登录+搜索+读表格+翻页，不创建/修改/删除任何工单

### M6 Excel 输出模块 (excel_writer.py)

**两份 Excel 输出**:

| 文件 | 内容 | 颜色标记 |
|------|------|----------|
| 漏单清单 | 有效邮件的字段提取+工单比对结果 | 红色=漏单、黄色=日期异常、蓝色=待确认、橙色=多匹配、灰色=不确定 |
| 过滤清单 | 被过滤的邮件及原因 | — |

**漏单清单字段**: 发件人邮箱、发件日期、主题、正文(精简)、代理、客户、项目、需求、是否已录单、工单日期、下单日期、匹配状态、查询时间戳

> 日期口径: `工单日期` = 邮件发来的日期（发件日期）；`下单日期` = 工单系统返回的下单日期（页面原始文本）。

## 工具模块

### fuzzy_match.py

- `normalize_text(text)`: OpenCC 简繁统一
- `fuzzy_match_pair(a, b, threshold)`: rapidfuzz 相似度匹配
- `fuzzy_search(text, candidates, high_threshold, medium_threshold)`: 模糊搜索

### attachment_parser.py

- **支持格式**: .xlsx .xls .pdf .docx .csv .txt .zip .rar .jpg .jpeg .png .bmp .tiff
- **OCR**: PaddleOCR `lang="ch"`，全程只初始化一次，失败后标记跳过
- **RAR 解压**: 三级降级 rarfile → patool → pyunpack
- **ZIP 解压**: 递归解析压缩包内所有支持的文件
- `extract_emails_from_text(text)`: 正则提取文本中的邮箱地址

## GUI 界面 (gui.py)

**PyQt5 界面，包含**:

- **时间范围选择器**: 开始日期 + 结束日期 (QDateEdit)
- **邮箱汇总表导入按钮**: 导入附件三 Excel (实时更新)
- **账号配置区**: 阿里邮箱地址、阿里邮箱密码、工单系统账号、工单系统密码 (密码框遮蔽)
- **开始/停止运行按钮**
- **进度显示**: `当前/总数` 格式 (非百分比)
- **日志面板**: 实时显示运行日志
- **查看结果按钮**: 运行完成后打开输出 Excel

**状态持久化** (session_state.json):

- 保存时机: 开始运行、导入汇总表、关闭窗口
- 恢复内容: 时间范围、邮箱汇总表路径

## 配置文件 (config.yaml)

`config.yaml` 已在 `.gitignore` 中忽略（含敏感凭据），下表列出全部 key 及用途。代码中未读取的 key 自动忽略。

### email — 阿里企业邮箱 IMAP

| Key | 类型 | 默认 | 说明 |
|-----|------|------|------|
| `imap_server` | str | `imap.qiye.aliyun.com` | IMAP SSL 服务器地址 |
| `imap_port` | int | `993` | IMAP SSL 端口 |
| `address` | str | — | 完整邮箱地址（GUI 也可改） |
| `password` | str | — | 客户端授权码（**非登录密码**），阿里邮箱后台生成 |
| `mailbox` | str | `INBOX` | 要读取的邮箱文件夹 |

### workorder — 工单系统 RPA 与查询缓存

| Key | 类型 | 默认 | 说明 |
|-----|------|------|------|
| `url` | str | `https://mp.ecopv-epr.com/workbench/main` | 工单系统登录页 |
| `username` / `password` | str | — | 工单系统账号密码（GUI 可改） |
| `timeout` | int | `30` | Playwright 页面操作默认超时（秒） |
| `query_interval_seconds` | float | `3` | 条目间间隔，防工单系统限流 |
| `query_cache_path` | str | `storage/query_cache.json` | 阶段二查询缓存落盘路径 |
| `query_cache_ttl_days` | int | `7` | 缓存有效期（`0`=永不过期） |
| `max_total_seconds` | int | `0` | 单次批量查询总超时秒数（`0`=不限）；超时强制终止，剩余条数下次续跑 |
| `selectors` | dict | `{}` | 选择器覆盖。**页面结构变化时只改这里不需改代码**。可用键：`login_account` / `login_password` / `nav` / `dropdown_option` / `company_input` / `query_button` / `table`，值为 CSS 选择器列表 |

### reference_tables — 参照表路径

| Key | 默认 | 说明 |
|-----|------|------|
| `internal_emails` | `data/internal_emails.xlsx` | 附件二：内部邮箱表（区分内/外部邮件） |
| `agent_emails` | `data/agent_emails.xlsx` | 附件三：代理邮箱对照表 |
| `project_names` | `data/project_names.xlsx` | 附件四：标准项目名称表 |

### fuzzy_match — 模糊匹配阈值

| Key | 默认 | 说明 |
|-----|------|------|
| `opencc_config` | `s2t.json` | OpenCC 简繁转换配置（`s2t`/`t2s`） |
| `high_threshold` | `90` | 高置信度（≥此值认为匹配） |
| `medium_threshold` | `80` | 中置信度（候选） |
| `low_threshold` | `80` | 模糊重查阈值 |

### workorder_match — 工单比对维度

| Key | 默认 | 说明 |
|-----|------|------|
| `date_tolerance_days` | `60` | 邮件日期 vs **平台下单日期**容差（`工单日期` 本身就是邮件日期，不参与比对） |
| `customer_threshold` | `85` | 客户名相似度阈值（rapidfuzz） |

### ocr — OCR 兜底开关

| Key | 类型 | 默认 | 说明 |
|-----|------|------|------|
| `fallback` | bool | `true` | 图片附件 OCR 是否作为字段提取的**兜底手段**。常规源（标题/正文/文档附件/文件名）+ LLM 之后仍缺**客户或项目**才触发（不含代理触发，避免代理表空时退化为常态）。图片持久化到 `cache/attachments/`，按内容 hash 去重，关闭后图片仅落盘不识别 |

### llm — DeepSeek 配置

| Key | 默认 | 说明 |
|-----|------|------|
| `api_key` | 空 | DeepSeek API Key；**为空则自动跳过 LLM**，仅用规则引擎 |
| `base_url` | `https://api.deepseek.com` | API 基础 URL |
| `model` | `deepseek-chat` | 模型名（不要用 `deepseek-v4-flash`） |
| `timeout` | `30` | 请求超时（秒） |
| `max_tokens` | `1000` | 单次响应最大 token |
| `temperature` | `0.1` | 采样温度（低温度更稳定） |

### output / logging — 输出与日志

| 段 | Key | 默认 | 说明 |
|----|-----|------|------|
| `output` | `dir` | `output` | 输出目录 |
| `output` | `missing_template` / `filtered_template` | `漏单清单_{timestamp}.xlsx` / `过滤清单_{timestamp}.xlsx` | 一站式兼容模板（v1.0 行为） |
| `logging` | `dir` | `logs` | 日志目录 |
| `logging` | `level` | `INFO` | 日志级别 |

## 数据依赖

| 参照表 | 文件 | 格式 | 说明 |
|--------|------|------|------|
| 附件二 | data/internal_emails.xlsx | 单列邮箱 | 公司内部邮箱地址，用于区分内外邮件 |
| 附件三 | data/agent_emails.xlsx | 代理名+简称+邮箱 | 代理邮箱对照表，邮箱→代理名映射 |
| 附件四 | data/project_names.xlsx | 项目编号+名称+国家+类型 | 标准项目名称，用于项目扫描和标准化 |

**当前状态**: 附件三通过 GUI 导入 (172 条)，附件二和附件四为空 (data/ 目录为空)，影响过滤和项目提取准确率。

## Python 依赖

```
PyQt5>=5.15
PyYAML>=6.0
openpyxl>=3.0
xlrd>=2.0
rapidfuzz>=3.0
opencc-python-reimplemented>=0.1
pdfplumber>=0.7
python-docx>=1.0
paddleocr>=2.0
rarfile>=4.0
patool>=4.0
pyunpack>=0.3
playwright>=1.40
```

## 已知问题与待办

1. ~~附件二/四为空~~ **已解决 (2026-09-10)**: GUI 新增"内部邮箱表导入/项目表导入"按钮；项目提取增加离线规则兜底（"国家+业务类型"相邻/连排/共享前缀模式），无附件四也能提取项目
2. **IMAP 超时**: **已缓解 (2026-09-10)**: 原始邮件按 `cache/mails/{mailbox}_{uid}.eml` 落盘缓存，重跑同日期范围直接命中不再拉网络；重连阈值只统计真实网络拉取（`email.reconnect_batch_size` 可配）；断连后已拉数据不丢
3. **工单系统页面适配**: **已缓解 (2026-09-10)**: 选择器全部抽到 `workorder.selectors` 配置段（默认值内置在 `workorder_checker.py` 的 `DEFAULT_SELECTORS`，config 只需覆盖变化的键）；新增 `probe_workorder.py` 探测脚本（登录后导出页面表单控件与表头结构到 output/probe_result.json，用于校准选择器）
4. ~~PaddleOCR 版本兼容~~ **已解决 (2026-09-10)**: mkldnn 在本机触发段错误 → paddle 导入前设 `FLAGS_use_mkldnn=0` + 构造参数 `enable_mkldnn=False`（2.x 回退环境变量）；返回格式同时兼容 2.x（`[[box,(text,score)]]`）与 3.x（`rec_texts` 字典）
5. **M5 串行污染 bug 已修 (2026-09-10)**: 旧版把所有行查询结果汇总成全池比对，A 行模糊重查（公司名前4字）返回的别家工单可能被 B 行"匹配成功"掩盖漏单。现改为每行只与本行查询结果比对（gui.py + search_batch 均已按行隔离）
6. **M4 优先级 bug 已修 (2026-09-10)**: `if proj_raw and "/" in proj_raw or "+" in proj_raw` 运算符优先级错误（空值时 `"+" in proj_raw` 仍会执行）。现所有非空项目统一走 normalizer 标准化
7. **LLM 模型名**: 使用 `deepseek-chat` (非 `deepseek-v4-flash`)
8. **OCR 改为懒加载兜底 (2026-09-10)**: 图片附件不再解析阶段就 OCR。常规解析阶段只持久化图片到 `cache/attachments/`（按内容 hash 去重，zip/rar 内图片同样处理）并登记 `ocr_pending`；字段提取 E3 阶段（规则+LLM 之后）仍缺**客户或项目**才触发 OCR 兜底（`ocr.fallback` 配置开关）。触发条件不含代理（代理靠邮箱表，表未命中时几乎每封都缺，单独触发会让兜底退化为常态；OCR 中若出现代理邮箱则机会性回填）。OCR 结果标记来源=`OCR图片兜底`、置信度强制 low（人工复核），多项目拆多行，找不到目标信息保持原空结果。附带修复：`_extract_company_from_attachment_text` 增加无标签裸公司名兜底（订单截图直接印公司名的场景）
9. **阶段一/二入口解耦 (2026-09-10)**: GUI 加运行模式选择（一站式 / 仅阶段一 / 仅阶段二）。WorkerThread.run() 拆为 `_run_stage1()`（邮件解析输出三份产出）/ `_run_stage2()`（读 stage1 输出 xlsx→前置校验→RPA→输出 workorder_check_result.xlsx）/ `_run_all()`（兼容原行为）。阶段二模式隐藏时间范围/邮箱导入等控件，新增"选择阶段一输出文件"按钮；session_state 持久化 mode + stage2_input_path
10. **阶段二断点续查 + 请求间隔 (2026-09-10)**: `WorkOrderChecker.search_one` 加 `query_cache.json` 缓存（按 代理-公司-项目 三联键），`storage/query_cache.json` 落盘，TTL 默认 7 天（`query_cache_ttl_days`），启动时命中直接复用不再走 RPA。`search_batch` 加条目间间隔 `query_interval_seconds`（默认 3s，防工单系统限流）+ `max_total_seconds` 总超时强制终止（剩余条数下次续跑）
11. **阶段二前置校验 (2026-09-10)**: `WorkOrderChecker.preprocess_rows()` 静态方法：①跳过「代理空 且 置信度=需人工确认」的条目（输出告警，不查 RPA）；②按 (公司,项目) 三联键去重（重复条目合并查询，节省 RPA 时间）。跳过的条目录入 workorder_check_result.xlsx 的"跳过后说明"列
12. **输出文件命名与分类 (2026-09-10)**: 默认将阶段一结果写入 `output/stage1_email/`，阶段二结果写入 `output/stage2_workorder/`，工作台导出写入 `output/manual_review/`；`output/diagnostics/` 作为探测文件和调试截图的统一预留目录。每类仍保留稳定名和时序副本，阶段二候选列表会递归扫描这些目录；旧版直接放在 `output/` 根目录的文件仍兼容读取。

## 测试

```bash
cd b:\TRAE_Project\6aa0bf7bc4ecce8d9c80e568\mail_audit_bot
python test_offline.py    # 离线全链路测试（纯构造数据，秒级，44 项断言）
python test_modules.py    # 在线集成测试（需真实邮箱/网络，约 5 分钟，部分断言基于旧OCR行为，已过时）
```

test_offline.py 覆盖: M2 过滤 / M3 规则提取(无附件四) / M4 拆分 / M5 按行比对(含串行污染回归) / M6 Excel / OCR 新旧格式兼容 / OCR 懒加载兜底(触发条件/来源标记/多项目拆行/空结果/开关) / **阶段二前置校验(空代理+重复合并)** / **阶段二查询缓存(命中/不命中/重启/过期)** / **阶段一二输出文件稳定名+时序副本**.

## Git 操作

```powershell
# 保存快照
cd b:\TRAE_Project\6aa0bf7bc4ecce8d9c80e568\mail_audit_bot
git add -A
git commit -m "描述改动内容"

# 查看历史
git log --oneline

# 回退到某个快照
git stash
git checkout <commit_id>

# 恢复到最新
git checkout master
```

## SOP 文档

原始 SOP 文档: `c:\Users\35144\.trae-cn\attachments\6aa0bf7bc4ecce8d9c80e56b\邮件&系统漏单审核SOP及自动化需求.docx`

SOP 定义了:
- 邮件识别规则 (方法一/方法二)
- 字段提取优先级: 名称 → 邮件正文 → 附件
- 项目拆分规则和示例
- 工单核对匹配条件: 日期+代理+客户+项目
- 日期规则: 邮件日期与平台下单日期相差不超过 2 个月（`工单日期` = 邮件发来日期，`下单日期` = 工单系统返回）
- 输出格式: 附件一模板
```
