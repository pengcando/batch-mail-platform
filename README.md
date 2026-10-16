# 邮件发送平台

面向批量邮件作业的网页平台。当前已上线第一个功能：**对账单批量生成与发送** —— 按城市 + 账期抽数生成每个商户的对账单 xlsx、渲染开票信息邮件正文，并批量发送。

架构是「外壳 + 功能插件」，后续新增功能只需各加一个后端包和一个前端模块，不动主框架。

> **首次使用前请做两处替换**（代码中已用占位符标出）：
> 1. `app/config.py` —— 改发件人邮箱与 SMTP 主机
> 2. `app/features/petty_cash/engine.py` —— 改 `SUBJECT_TEMPLATE` / `ATTACH_TEMPLATE` 里的「公司名」前缀

---

## 1. 快速开始

### 本机运行

```bash
# 方式一：双击 run.bat（自动建 .venv、装依赖、开浏览器）

# 方式二：手动
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
```

打开 <http://127.0.0.1:8000>。

### 部署到服务器

单进程即可，无需数据库、无需本机 Excel：

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
```

建议通过环境变量收敛行为（见第 6 节）。若要放 Nginx 后面，注意 **上传体积**（默认单文件 60MB）与 **任务超时**（大批量发信可能跑几分钟，轮询接口本身很快）。

---

## 2. 目录结构

```
mail-platform/
├─ app/
│  ├─ main.py                    应用入口：注册功能路由 + 挂载静态页
│  ├─ config.py                  全局配置（可被环境变量覆盖）
│  ├─ core/                      与具体业务无关的公共能力
│  │  ├─ jobs.py                 内存任务仓库（进度/结果/取消）
│  │  ├─ mailer.py               SMTP 发信（465/SSL）
│  │  ├─ email_html.py           邮件正文渲染（金额两位小数、空行与缩进保留）
│  │  └─ xlsx_style.py           跨工作簿搬运样式 / 行重排
│  └─ features/
│     ├─ __init__.py             功能注册表
│     └─ petty_cash/             功能一：对账单批量作业
│        ├─ engine.py            业务引擎（列映射、业务公式、工作簿生成）
│        ├─ pipeline.py          阶段 A→B→C 编排 + 进度回调
│        └─ routes.py            HTTP 接口
├─ web/                          纯静态前端（无构建步骤、无 CDN）
│  ├─ index.html
│  ├─ css/app.css
│  └─ js/app.js                  应用外壳：路由 / 功能注册 / 弹窗
│     js/features/petty-cash.js  功能一的前端界面
├─ storage/                      运行时产物（jobs/*.zip、uploads/）
├─ requirements.txt
└─ run.bat
```

---

## 3. 功能一：对账单批量作业

### 输入

| 项 | 说明 |
|---|---|
| ① 不开票商户明细-{城市}.xlsx | 商户名单（`M`），提供收件人 / 抄送 / 进度列 |
| ② 商户对账工具-开票情况-V2.0.xlsx | 抽数工具（`T`），含城市明细页、对账单模板页、增减项页、开票信息页 |
| 城市 | 下拉单选，选项由②动态解析出来，与城市工作表名严格一致 |
| 账期年月 | `YYYY-MM`，用于文件名与邮件主题 |
| 邮箱授权码 | 发件人与 SMTP 主机在 `app/config.py` 里配置（默认 `sender@example.com` / `smtp.example.com:465` SSL），**用前请改成你自己的** |
| 开关 | 预览模式（默认开）、跳过已处理行（默认开）、整表数值格式、月份列取值 |

### 处理流程

| 阶段 | 实现位置 |
|---|---|
| A5–A13 打开文件、清筛选、按 A 列降序、读取待处理商户并记录绝对行号 | `pipeline.run` 前半段 + `xlsx_style.reorder_rows_by_key` |
| B1–B8 按商户取数（D 列筛选，得到行数 `n`） | `pipeline.run` 的 `buckets` 分桶 |
| B8a/B8b `n=0` → 写「本月无流水」、跳过本轮 | `pipeline.run` 分支 1 |
| B8c–B8f `n>0` → 粘贴映射（城市 A–R→账单 C–T、S→W，避开 U/V 计算列） | `engine.build_statement_row` |
| B9–B16 生成对账单文件、只留「表头 + n 行」、追加 `增减项` 页 | `engine.build_statement_workbook` |
| B17–B23 增减项（`m=0` 只留表头；`m>0` 贴数据） | `engine.build_adjust_rows` + `pipeline` |
| B24–B28 开票信息 C10 置 0.00、读取 A1:I10、渲染 HTML | `engine.build_invoice_context` + `core/email_html.py` |
| B29–B35 组装并发送邮件（主题/正文/附件） | `pipeline.run` + `core/mailer.send_mail` |
| B36–B38 回写 `M` 表 G 列状态 | `pipeline.run` 尾部 |
| C1–C5 收尾 | `pipeline.run` 结尾 |

### 输出（ZIP 打包下载）

- N 个 `公司名&{商户}对账单-{年月}.xlsx`（每个含 `对账单` + `增减项` 两张表）
  > 「公司名」是占位符，请到 `app/features/petty_cash/engine.py` 的
  > `SUBJECT_TEMPLATE` / `ATTACH_TEMPLATE` 换成你自己的品牌前缀。
- 回写状态后的 `不开票商户明细-{城市}.xlsx`（**仅正式发送模式包含**）
- `发信结果-{城市}-{年月}.csv`（商户 / 收件人 / 抄送 / 状态 / 说明，UTF-8 BOM）
- `_邮件正文预览.html`（第一封邮件的正文，也可在页面上直接看）

状态取值：`已完成` / `本月无流水` / `未完成`。

### 运行时的实时进度条

进度条位于**运行结果卡片最下方并吸底**（`position: sticky; bottom: 0`），
所以边往下翻逐商户结果、边能一直看到进度。运行时显示：

- 当前阶段 + 正在处理的商户名（服务端 `job.current`）
- 百分比、已完成 / 总数
- 已运行时间、按已完成样本线性外推的**预估剩余**、处理速率（商户/秒）

实现要点：进度条 DOM **只创建一次**，之后由 `renderProgress()` 原地改写文本与
`bar.style.width`，因此 CSS 的 `width` 过渡真正生效（旧写法每 700ms 重建
innerHTML，宽度只能跳变）。总数未知或首个商户未跑完时切换为 `indeterminate`
流动条纹态；结束后停在全绿/全红并定格。

---

## 4. 关键实现决策

原本的流程是「本机 Excel + 宏/RPA」操作，搬到服务端必须把 Excel 的公式求值替换掉。以下三点是**有意为之**的实现决策：

1. **对账单页不再带公式。** 模板里 `对账单` 页的单元格全是跨表引用（如 `='账单数据-原始数据'!Y`），线上化后由引擎按列映射直接算成字面值写入。好处是产物不含跨表引用，永远不会出现失效的 `#REF!`；代价是它不再随源文件联动（本来就是发给客户的静态账单，无影响）。

2. **业务公式在 Python 里复刻。**
   - 服务费开票金额 = `IF(AND(商户开票情况="是", 客户开票情况="否"), 调整后销售额-调整后结算额, 0)`
   - 客户发票开票方 = `IF(客户开票情况="否", "贵司", "公司名")`
   - 邮件里的「本期销售额/结算额」= 该商户全部明细行的 `调整后销售额` / `调整后结算额` 之和，服务费金额 = 两者之差。

   > 若你的业务口径不同，改 `engine.py` 里的 `_invoice_party()` 与 `build_statement_row()` 即可。

3. **样式从模板快照搬运**（`core/xlsx_style.py`），保证出品对账单与模板一致。

以下几处默认行为可按你的业务调整：

| 项 | 当前实现 | 说明 |
|---|---|---|
| `增减项` 页内容 | **只保留该商户的行**（表头 + 匹配行） | 这是发给客户的账单，把别的商户的增减项一并贴过去属于数据泄露 |
| `对账单` 的「月份」列 | 默认取运行日期往前一个月 | 界面上可切到「直接取本次账期」 |
| 「整表设为数值格式」 | **默认关闭** | 整表套数值格式会把「月份」列变成日期序列号、把佣金率列放大 100 倍 |

---

## 5. 已处理的源文件坑点

| 坑 | 处理 |
|---|---|
| `增减项` 表被格式化到 XFD 列，`UsedRange` 虚增 | 只按实际有值的列（A–F）读取，绝不使用整个已用区域 |
| E/F 列多邮箱混用中英文分号 | `engine.split_emails` 按 `; ； , ，` 与空白切分，清洗后逐个校验格式并去重 |
| `M` 表工作表名固定为 `Sheet1` | 固定按 `Sheet1` 取，取不到直接报错，不静默 fallback |
| `开票信息!A1` 是带换行的长文本 | 按 `\n` 拆成多个 `<p>`，空行与 4 空格缩进用 `&nbsp;` 严格保留 |
| 城市明细页按商户取数需全表筛选 | 一次扫描按商户分桶，避免每轮重复遍历整页 |
| 重跑时重复发信 | 「跳过已处理行」默认开启，只处理 `G` 列为空的行 |
| 全部被跳过导致"假成功" | 空结果显式提示原因、不打空 zip，下载按钮禁用 |

---

## 6. 环境变量

| 变量 | 默认 | 说明 |
|---|---|---|
| `MAIL_SENDER` | `sender@example.com` | 发件人，**请改成你自己的** |
| `SMTP_HOST` | `smtp.example.com` | SMTP 主机，**请改成你自己的**（如 `smtp.exmail.qq.com` / `smtp.qq.com`） |
| `SMTP_PORT` | `465` | 端口 |
| `SMTP_USE_SSL` | `1` | 是否 SSL |
| `ALLOW_REAL_SEND` | `1` | 设为 `0` 彻底锁死「正式发送」，服务端只允许预览 |
| `MAX_MERCHANTS` | `0` | 单次最多处理商户数，`0` 为不限制；上线初期建议设个上限 |

---

## 7. 新增一个功能

后端：

1. 新建 `app/features/<your_feature>/`，写 `engine.py` / `pipeline.py` / `routes.py`，`routes.py` 里暴露 `router`；
2. 包的 `__init__.py` 里写一个 `FEATURE = {"id": ..., "name": ..., "desc": ..., "icon": ...}`；
3. 把包路径加进 `app/features/__init__.py` 的 `_MODULES`。

前端：

1. 新建 `web/js/features/<your_feature>.js`，调用 `window.WB.register({ id, name, desc, icon, mount })`；
2. 在 `web/index.html` 底部加一行 `<script>`。

导航、路由、弹窗、API 封装都由外壳统一提供。

---

## 8. 验证方式

本仓库不含测试脚本与样本数据，请用你自己的业务文件按下面顺序验证：

1. **离线校验**：直接用网页端跑一次「预览模式」，检查产出的 zip
   （应为 N 份对账单 + 回写后的名单 + 发信结果 CSV + 邮件正文预览 HTML）。
2. **界面校验**：确认进度条在运行中实时走动、结束后定格。
3. **正式发送**：先用少量商户验证收信与附件版式，再批量执行。

> ⚠️ 网页端默认就是 **预览模式**，不发任何邮件。正式发送需在高级设置里显式关闭预览、
> 填入邮箱授权码，并经过二次确认弹窗。
