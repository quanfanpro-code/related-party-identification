# 关联方维度自适应分批导出实施计划

> **执行约束：** 未授权子任务，按 `executing-plans.md` 在当前任务中串行实施；每个步骤用复选框跟踪。用户未授权 Git 写入，因此本计划不提交、不合并、不建立工作树，仅保留可审阅的本地变更。

**目标：** 将注协行业库导出从一次请求全部 61 个维度，改为仅查询关联方核查所需的 20 个维度，并按“10+10，失败组逐级拆分，成功维度不重下”的规则可靠完成四川江擎航空科技有限公司本次查询。

**总体方案：** 复用现有企业名单上传、批量查询、下载中心任务关联和 ZIP 安全校验能力；企业名单只上传一次，20 个维度先分成两组，每次只允许一个下载任务在途。每组经两个完整轮询窗口仍失败时才拆分为更小组；每批有效文件立即校验、登记并合并到平铺成果目录，后续只重试未完成维度。

**技术栈：** Python 3.9+、标准库、requests、openpyxl、unittest/pytest、Windows、注协行业库现有 HTTP 接口。

## 全局约束

- 固定 20 个维度，不再默认获取和导出全部 61 个维度。
- 首轮固定为 10+10；失败组按 10→5+5、5→3+2、3→1+1+1、2→1+1 拆分，即严格执行用户确认的“10→5+5→3+2→单维度”，且不会产生空组。
- 已通过校验或已明确无数据的维度永久跳过，不重复下载。
- 一个维度只有在导出文件存在且校验为空时才能记为“明确无数据”；静默缺失仍视为未完成。
- 单维度经过两个完整轮询窗口仍失败时暂停并保留状态，不进入关联方候选扩展。
- 普通请求前随机等待 2–5 秒；轮询间隔随机等待 10–20 秒；成功批次之间随机等待 15–30 秒。测试中注入空等待函数，不真实休眠。
- 同一时刻最多一个注协下载任务在途，避免下载中心任务串线。
- 所有状态写入使用原子替换并对 Windows 临时占用做有限重试；不得在状态中保存 Cookie、密码或其他认证秘密。
- 批次 ZIP 和有效 Excel 均保留；成果目录继续向报告引擎提供平铺 Excel 文件，不改变后续读取方式。
- 不新增第三方依赖，不删除旧的 61 维任务和已有成果。
- 修改任何现有文件前，先备份到 `C:\Users\27651\BackUp\related-party-identification_时间戳\`，再比对 SHA-256；同一文件只串行修改。
- 包含中文的新增或改写文本使用 UTF-8 with BOM；状态 JSON 按既有安全约定继续使用无 BOM UTF-8。
- 每项功能遵循 RED（先看到测试失败）→ GREEN（最小实现）→ 回归验证。

---

## 文件职责图

- 新建 `scripts/state_io.py`：只负责带有限重试的原子 JSON 状态写入，供任务状态和导出状态共同复用。
- 修改 `scripts/cicpa/exporter.py`：固定维度目录、分组状态机、单任务串行触发、分批下载校验、失败拆分、成果合并和恢复执行。
- 修改 `scripts/related_party_workflow.py`：接入自适应导出器，并在 20 个维度闭环前阻断候选扩展和报告生成。
- 修改 `scripts/validate_bundle.py`：忽略 `.pytest_cache`、`__pycache__` 等标准缓存目录，避免把测试缓存误判为交付文件。
- 修改 `scripts/tests/test_exporter.py`：覆盖维度目录、10+10、拆分、成功跳过、部分 ZIP、明确无数据、文件冲突、等待区间和恢复。
- 新建 `scripts/tests/test_state_io.py`：覆盖 Windows 原子替换临时占用的有限重试。
- 修改 `scripts/tests/test_workflow.py`：覆盖工作流关口、旧状态兼容和完整成果后再进入候选扩展。
- 修改 `SKILL.md`、`README.md`、`CONTEXT.md`、`references/user-flow.md`、`references/dimensions.md`、`NOTICE`、`references/SOURCES.json`、`agents/openai.yaml`、`test-prompts.json`：同步实际流程、用户话术、来源和验收提示。

---

### 任务 1：执行前安全快照与基线固定

**文件：**

- 只读：计划列出的全部现有文件
- 创建：`C:\Users\27651\BackUp\related-party-identification_YYYYMMDD_HHMMSS\` 下的同名备份文件

**产出：** 每个拟修改现有文件都有修改前副本和一致的 SHA-256；测试基线被记录，不触碰功能代码。

- [ ] **步骤 1：创建唯一备份目录**

  使用已显式设置 UTF-8 输出的 PowerShell 创建时间戳目录，不在项目目录内创建备份。

- [ ] **步骤 2：逐一复制现有文件**

  复制 `scripts/cicpa/exporter.py`、`scripts/related_party_workflow.py`、`scripts/validate_bundle.py`、`scripts/tests/test_exporter.py`、`scripts/tests/test_workflow.py`、`SKILL.md`、`README.md`、`CONTEXT.md`、`references/user-flow.md`、`references/dimensions.md`、`NOTICE`、`references/SOURCES.json`、`agents/openai.yaml`、`test-prompts.json`。本轮已备份的 `CONTEXT.md` 仍再次纳入统一实施快照，便于整轮回滚。

- [ ] **步骤 3：验证备份**

  对源文件和备份文件分别执行 SHA-256，要求同名文件哈希完全一致；若任一文件不一致，停止实施。

- [ ] **步骤 4：重跑现有基线**

  运行：

  ```text
  python -X utf8 -m pytest -q -p no:cacheprovider
  ```

  预期：保留当前已知基线“101 passed、1 failed、48 subtests passed”；唯一失败仍是校验器误扫描 `.pytest_cache/README.md`，没有新增失败。

---

### 任务 2：修复交付校验器误扫描缓存目录

**文件：**

- 修改：`scripts/validate_bundle.py`
- 修改：`scripts/tests/test_workflow.py`

**接口：**

- 产出：`_text_files(root: Path)` 不返回路径任一部分属于 `{'.git', '.pytest_cache', '__pycache__'}` 的文件。
- 后续依赖：最终包校验可以在不删除缓存的情况下通过。

- [ ] **步骤 1：添加失败测试**

  在 `BundleValidationTests` 中创建含 `.pytest_cache/README.md`（无 BOM）和正常中文 Markdown（有 BOM）的临时技能包，断言 `validate_bundle(...).checks['encoding']` 为真；同时创建普通目录中的无 BOM Markdown，断言仍会失败，防止放宽实际编码规则。

- [ ] **步骤 2：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "缓存目录" -q -p no:cacheprovider
  ```

  预期：缓存目录测试失败，错误指向 `.pytest_cache/README.md`。

- [ ] **步骤 3：作最小修改**

  在 `scripts/validate_bundle.py` 定义：

  ```python
  IGNORED_DIRECTORY_NAMES = {".git", ".pytest_cache", "__pycache__"}
  ```

  `_text_files()` 在文件路径任一目录命中该集合时跳过；不更改普通文件的 BOM 和 UTF-8 规则。

- [ ] **步骤 4：验证局部与全量基线**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "缓存目录" -q -p no:cacheprovider
  python -X utf8 -m pytest -q -p no:cacheprovider
  ```

  预期：新增测试通过；原先唯一失败消失，至少 102 个测试通过且 48 个子测试通过。

---

### 任务 3：统一修复 Windows 状态文件偶发占用

**文件：**

- 创建：`scripts/state_io.py`
- 创建：`scripts/tests/test_state_io.py`
- 修改：`scripts/cicpa/exporter.py`
- 修改：`scripts/related_party_workflow.py`

**接口：**

- 产出：

  ```python
  def atomic_write_json(
      path: Path,
      payload: dict,
      *,
      replace_attempts: int = 3,
      replace_delay_seconds: float = 0.2,
      replace: Callable[[Path, Path], None] = os.replace,
      sleeper: Callable[[float], None] = time.sleep,
  ) -> None
  ```

- 后续依赖：`save_export_state()`、`save_task_state()` 都通过该函数落盘。

- [ ] **步骤 1：添加失败测试**

  在 `test_state_io.py` 注入一个前两次抛出 `PermissionError`、第三次成功的 `replace`，断言调用三次、目标 JSON 可读且不带 BOM；再注入连续三次失败，断言第三次后原异常被抛出且旧目标文件仍可读取。

- [ ] **步骤 2：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_state_io.py -q -p no:cacheprovider
  ```

  预期：因 `scripts.state_io` 尚不存在而失败。

- [ ] **步骤 3：实现一个共用写入函数**

  使用同目录 `.tmp` 文件、`flush()`、`os.fsync()` 和 `os.replace()`；只捕获 `PermissionError`，最多尝试三次，前两次分别等待 0.2 秒；失败后不改动既有目标文件，不吞掉异常。

- [ ] **步骤 4：替换两处重复实现**

  `save_export_state()` 调用 `atomic_write_json(path, asdict(state))`；`save_task_state()` 调用 `atomic_write_json(path, asdict(state))`。删除两处重复的临时文件写入代码，保持状态 JSON 无 BOM。

- [ ] **步骤 5：验证状态读写与安全字段**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_state_io.py scripts/tests/test_exporter.py -k "状态 or atomic" -q -p no:cacheprovider
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "状态" -q -p no:cacheprovider
  ```

  预期：有限重试、旧文件保护、无 BOM、无 Cookie/密码字段全部通过。

---

### 任务 4：建立固定 20 维目录和可恢复的分组状态机

**文件：**

- 修改：`scripts/cicpa/exporter.py`
- 修改：`scripts/tests/test_exporter.py`

**接口：**

- 产出常量：

  ```python
  REQUIRED_DIMENSIONS = (
      ("S0000002", "基础工商信息"),
      ("S0000006", "股东信息"),
      ("S0000103", "最新公示股东"),
      ("S0000032", "实际控制人"),
      ("S0000020", "最终受益人"),
      ("S0000037", "主要人员（高管）"),
      ("S0000123", "核心团队"),
      ("S0000104", "对外投资（新）"),
      ("S0000105", "参控股企业"),
      ("S0000107", "发票信息"),
      ("S0000119", "客户"),
      ("S0000118", "供应商"),
      ("S0000016", "变更记录"),
      ("S0000019", "法定代表人变更"),
      ("S0000018", "经营异常"),
      ("S0000013", "股权质押"),
      ("S0000012", "动产抵押"),
      ("S0000041", "商标"),
      ("S0000036", "软件著作权"),
      ("S0000101", "微信公众号"),
  )
  ```

- 扩展 `ExportState`：`required_dimensions`、`dimension_results`、`pending_groups`、`active_group`、`batch_history`、`poll_windows`；旧 JSON 缺少字段时由 dataclass 默认值兼容加载。
- 产出：`split_dimension_group(group: List[str]) -> List[List[str]]`，保持原顺序；10 维拆成 5+5，5 维拆成 3+2，2–3 维直接拆成单维度。
- 产出：`next_dimension_group(state: ExportState) -> List[str]`，过滤状态为 `completed` 或 `no_data` 的代码。

- [ ] **步骤 1：添加目录和拆分失败测试**

  断言 20 个代码及中文名称与确认清单逐项一致且无重复；断言 10 拆为 5+5、5 拆为 3+2、3 拆为 1+1+1、2 拆为 1+1、1 不再拆。

- [ ] **步骤 2：添加恢复和成功跳过失败测试**

  构造首组已有 7 个 `completed`、1 个 `no_data`、2 个 `pending` 的状态，断言下一请求只含那 2 个未完成代码；保存再加载后结果一致。

- [ ] **步骤 3：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -k "二十维 or 拆分 or 成功维度" -q -p no:cacheprovider
  ```

  预期：常量、字段或函数尚不存在而失败。

- [ ] **步骤 4：实现最小状态模型**

  `start_export()` 仍只上传一次企业名单并触发一次批量查询，但不再把实时接口返回的 61 个代码直接交给导出接口；它先读取实时维度目录，逐项校验固定 20 个代码均存在且名称可对应，然后初始化两个各 10 个代码的 `pending_groups`。

- [ ] **步骤 5：实现过滤和拆分**

  拆分只处理失败组；所有取下一组的路径都先过滤 `dimension_results` 中的 `completed` 和 `no_data`，过滤后为空则直接取后续组，不触发空任务。

- [ ] **步骤 6：验证状态机**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -k "二十维 or 拆分 or 成功维度 or 恢复" -q -p no:cacheprovider
  ```

  预期：所有目录、拆分、跳过和恢复测试通过。

---

### 任务 5：实现单任务串行触发、人工节奏和两窗口失败判定

**文件：**

- 修改：`scripts/cicpa/exporter.py`
- 修改：`scripts/tests/test_exporter.py`

**接口：**

- `CicpaExporter.__init__()` 新增可注入参数：

  ```python
  sleeper: Callable[[float], None] = time.sleep
  randint: Callable[[int, int], int] = random.randint
  ```

- 产出：`trigger_next_group(state, state_path) -> bool`；有待办时触发一个任务并返回真，无待办时返回假。
- 产出：`wait_current_group(state, state_path, max_polls=18) -> Optional[Dict[str, Any]]`；一个窗口内只跟踪 `active_group` 对应任务。
- 产出：`handle_poll_window_failure(state, state_path) -> None`；第一个窗口只保留原组，第二个窗口才拆分；单维度第二窗口失败则标记暂停。

- [ ] **步骤 1：添加人工节奏失败测试**

  注入记录型 `sleeper` 和确定型 `randint`，断言触发请求前使用 2–5 秒区间、未完成轮询之间使用 10–20 秒区间、成功批次后且仍有待办时使用 15–30 秒区间；最后一批完成后不再额外等待。

- [ ] **步骤 2：添加两窗口与单任务失败测试**

  断言第一个轮询窗口结束后原 10 维组不拆；第二个窗口结束后才进入两个 5 维组；`active_group` 或 `task_id` 非空时再次调用触发函数不会发第二个请求；单维度第二窗口失败后状态为 `paused`，待办和历史仍在。

- [ ] **步骤 3：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -k "等待区间 or 两个窗口 or 单任务" -q -p no:cacheprovider
  ```

  预期：注入参数和新方法尚不存在而失败。

- [ ] **步骤 4：实现串行触发**

  每次触发前刷新下载中心任务号快照，将本组代码写入 `active_group`，清空上组 `task_id/task_name`，设置 `triggered_at`，调用既有 `_trigger_export(batch_no, active_group)`；写状态成功后才进入轮询。

- [ ] **步骤 5：实现等待和失败判定**

  每个轮询窗口最多 18 次；任务完成立即返回，任务尚未完成且还有轮询次数时按 10–20 秒休眠。窗口结束增加 `poll_windows`；第二个窗口结束才把失败组拆到队首，清空活动任务字段并保留历史。单维度无法再拆时写明代码、名称和两窗口结果，暂停而非继续。

- [ ] **步骤 6：验证触发与轮询**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -k "等待区间 or 两个窗口 or 单任务 or 任务关联" -q -p no:cacheprovider
  ```

  预期：所有区间边界、两窗口、单任务和原任务关联测试通过。

---

### 任务 6：实现每批 ZIP 校验、部分成功留存和无覆盖合并

**文件：**

- 修改：`scripts/cicpa/exporter.py`
- 修改：`scripts/tests/test_exporter.py`

**接口：**

- 产出：`download_current_group(state, task, state_path) -> List[str]`，返回本批新完成或明确无数据的维度代码。
- 批次归档：`dimension-batches/batch-001.zip`、`dimension-batches/batch-001_files/`，序号按 `batch_history` 递增。
- `dimension_results[code]` 的终态只能是 `completed` 或 `no_data`；未出现、损坏或冲突的维度仍留在待办队列。
- `batch_history` 记录请求代码、任务号、ZIP 路径、成功代码、无数据代码、缺失代码、结果和时间，不保存认证信息。

- [ ] **步骤 1：扩展测试 ZIP 构造器**

  让测试按“中文维度名→数据行”生成多个 Excel；空维度仍生成只有表头的工作簿；部分 ZIP 刻意省略指定工作簿。

- [ ] **步骤 2：添加完整、空数据和部分 ZIP 失败测试**

  断言完整 10 维 ZIP 全部记为 `completed`；只有表头的工作簿记为 `no_data`；只含 7 个有效文件的 ZIP 保留这 7 个结果，仅把缺失 3 维重新排队，下一请求不含已成功 7 维。

- [ ] **步骤 3：添加企业集合、路径穿越和哈希冲突测试**

  断言任一非空工作簿中的被查企业集合不符时该维度不完成；ZIP 路径穿越仍在写出前阻断；目标目录已有同名同哈希文件时复用，已有同名不同哈希文件时暂停并保留两个批次证据，绝不覆盖。

- [ ] **步骤 4：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -k "部分ZIP or 明确无数据 or 哈希冲突 or 路径穿越" -q -p no:cacheprovider
  ```

  预期：分批归档和逐维结果尚未实现而失败；既有路径穿越测试保持通过。

- [ ] **步骤 5：实现逐维识别和校验**

  用固定中文名称匹配 `.xlsx` 文件名；工作簿存在且只有表头时记为 `no_data`，存在数据时校验被查企业集合后记为 `completed`。ZIP 中缺失的预期名称只记为缺失，不把整批已验证文件作废。

- [ ] **步骤 6：实现安全合并**

  先安全解压到独立批次目录；对每个有效 Excel 计算 SHA-256。根成果目录无同名文件时复制，有同名同哈希时复用，有同名不同哈希时设置 `paused` 并写冲突说明；任何路径都不直接覆盖旧证据。

- [ ] **步骤 7：实现批次闭环**

  写完本批历史后，清空活动任务字段和窗口计数；若有缺失维度，将缺失维度作为一个失败组进入待办队首并遵循后续两窗口/拆分规则；若本批全部完成，按 15–30 秒节奏进入下一组。

- [ ] **步骤 8：验证分批文件处理**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_exporter.py -q -p no:cacheprovider
  ```

  预期：原下载安全测试和新增分批、部分成功、空数据、冲突测试全部通过。

---

### 任务 7：把自适应导出闭环接入关联方工作流

**文件：**

- 修改：`scripts/cicpa/exporter.py`
- 修改：`scripts/related_party_workflow.py`
- 修改：`scripts/tests/test_workflow.py`

**接口：**

- 产出：`run_to_completion(state, state_path, max_polls=18) -> Path`，串行驱动“取组→触发→等待→下载→登记→下一组”，仅在 20 个维度全部终态时返回成果目录。
- 修改：`_ensure_export(...) -> Path` 负责新建或加载状态，再调用 `run_to_completion()`。
- 关口：只有 `required_dimensions` 中每个代码均为 `completed/no_data`，工作流才能进入 `discovering`；否则保持 `waiting_export` 或 `paused_export`。

- [ ] **步骤 1：添加闭环失败测试**

  用假导出器模拟两组 10 维均成功，断言企业名单只上传一次、请求恰好两批、完成后才调用候选发现；模拟第二组失败拆分为 5+5，其中首个 5 成功，断言后续不重下该 5 维。

- [ ] **步骤 2：添加暂停和旧状态兼容失败测试**

  模拟单维度两窗口失败，断言候选发现和报告写入均未调用、任务状态含暂停原因和可恢复导出状态路径。加载旧版 `ExportState` 时，若状态已 `completed` 且成果目录存在则继续复用；旧版 `waiting/failed` 的 61 维任务不在线续跑，另起新 20 维状态文件且旧文件原样保留。

- [ ] **步骤 3：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "自适应导出 or 二十维关口 or 旧导出状态" -q -p no:cacheprovider
  ```

  预期：闭环方法和关口尚不存在而失败。

- [ ] **步骤 4：实现导出器闭环**

  `run_to_completion()` 每次循环都先保存状态，且任何异常都保留当前批次、完成维度和待办组；若状态为 `paused`，抛出含明确中文原因的 `ExportError`，不自行绕过。

- [ ] **步骤 5：实现工作流关口**

  `_ensure_export()` 对新任务创建新状态；对新版未完成状态直接恢复；对旧版 61 维未完成状态改用新的 `*-adaptive-export-state.json`，不覆盖旧 JSON 或旧 ZIP。20 维未闭环时设置任务阶段为 `waiting_export/paused_export`，不调用候选扩展。

- [ ] **步骤 6：验证工作流回归**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -q -p no:cacheprovider
  ```

  预期：新关口测试及既有输入、认证、发现、报告测试全部通过。

---

### 任务 8：同步整个 skill 的说明、提示词和来源记录

**文件：**

- 修改：`SKILL.md`
- 修改：`README.md`
- 修改：`CONTEXT.md`
- 修改：`references/user-flow.md`
- 修改：`references/dimensions.md`
- 修改：`NOTICE`
- 修改：`references/SOURCES.json`
- 修改：`agents/openai.yaml`
- 修改：`test-prompts.json`
- 修改：`scripts/tests/test_workflow.py`

**产出：** 所有面向用户和代理的流程都只描述固定 20 维、自适应拆分、人工节奏、完成关口和恢复规则，不再指导一次打包 61 维。

- [ ] **步骤 1：添加文档契约失败测试**

  在 `test_workflow.py` 读取上述文件，断言关键文件均包含“20 个维度”“10+10”“成功维度不重复下载”“两个完整轮询窗口”“单维度失败暂停”等核心语义；断言运行说明中不再存在“导出全部维度”“一次性 61 维”等旧指令。

- [ ] **步骤 2：确认测试先失败**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "分批导出文档契约" -q -p no:cacheprovider
  ```

  预期：现有文档仍含旧流程，测试失败。

- [ ] **步骤 3：局部更新运行说明**

  `SKILL.md` 和 `references/user-flow.md` 写清：先 10+10，失败组才拆，成功维度不重下；等待期间向用户报告当前批次和已完成数量；20 维闭环后才进入候选扩展。

- [ ] **步骤 4：更新维度和业务上下文**

  `references/dimensions.md` 列出 20 个代码、名称、关联方用途及“完成/明确无数据”判定；`CONTEXT.md` 保留已经批准的术语并补齐与状态字段的一致表述。

- [ ] **步骤 5：更新版本说明和来源**

  `README.md` 追加 2026-08-06 本轮问题、方案、未删除旧任务、验证结果和四川江擎实测结果；`NOTICE`、`references/SOURCES.json` 明确仍复用上游接口思路，但自适应调度、逐维状态和校验为本地实现，不虚构上游已有此能力。

- [ ] **步骤 6：更新代理入口和验收提示**

  `agents/openai.yaml` 和 `test-prompts.json` 增加分批失败、恢复、单维暂停和四川江擎查询的验收场景，确保新对话不会重新走 61 维一次打包。

- [ ] **步骤 7：验证文档契约和中文编码**

  运行：

  ```text
  python -X utf8 -m pytest scripts/tests/test_workflow.py -k "分批导出文档契约" -q -p no:cacheprovider
  python -X utf8 scripts/validate_bundle.py
  ```

  预期：文档契约通过；技能包编码、来源、许可、测试和结构全部有效。

---

### 任务 9：四川江擎实测、报告生成与整体验收

**文件：**

- 运行时创建：`C:\Users\27651\Desktop\四川江擎航空科技有限公司_关联方核查_运行时间\`
- 保留不改：`C:\Users\27651\Desktop\四川江擎航空科技有限公司_关联方核查_20260806_214433\` 及其旧 61 维任务证据
- 只读验证：项目全部已改文件和新文件

**完成标准：** 20 个维度全部为 `completed/no_data`，状态历史可说明每次请求和拆分；候选扩展只在闭环后运行；生成可交付的关联方候选报告和证据目录。

- [ ] **步骤 1：执行全部自动测试**

  运行：

  ```text
  python -X utf8 -m pytest -q -p no:cacheprovider
  ```

  预期：至少 102 个原有/基线测试加本轮新增测试全部通过，48 个子测试继续通过；不得只报告局部测试。

- [ ] **步骤 2：执行包校验和差异检查**

  运行：

  ```text
  python -X utf8 scripts/validate_bundle.py
  git diff --check
  git status --short
  ```

  预期：技能包有效、无空白错误；状态只显示计划内文件。任何意外文件立即停下复核，不自行删除。

- [ ] **步骤 3：验证编码和异常字符**

  检查所有新增/修改中文文本可按 UTF-8 解码，要求文本文件 BOM 符合技能包策略；统计 Unicode 替换字符为 0，异常连续问号为 0。状态 JSON 继续验证为无 BOM UTF-8。

- [ ] **步骤 4：确认登录后启动四川江擎新导出**

  使用现有 Firefox 登录状态；新建独立成果目录，企业名称为“四川江擎航空科技有限公司”。先请求第一批 10 维，验证成功后按 15–30 秒节奏请求第二批；仅当某组真实失败且经过两个完整窗口才执行拆分。

- [ ] **步骤 5：核验实测状态和文件**

  状态文件应显示 20 个固定代码全部为 `completed/no_data`，无其他 41 个维度；对照 `batch_history` 核验每批任务号、ZIP、文件哈希、缺失/空数据判定和等待窗口。随机抽查基础工商、股东/实际控制、高管、投资、客户/供应商、变更及知识产权类文件可打开且企业名称一致。

- [ ] **步骤 6：进入候选扩展并生成报告**

  只有步骤 5 通过后运行后续关联方候选发现和报告生成；报告区分“注协行业库证据”“公开资料佐证”“仅候选待核实”，不得把公开信息冲突直接写成已确认关联方。

- [ ] **步骤 7：最终交付核对**

  向用户报告：实际用了几批、每批维度数、是否触发拆分、20 维完成/无数据数量、报告与证据目录、自动测试和包校验结果、保留的旧任务位置。若单维度仍失败，则报告暂停维度和证据，不虚报完成。

---

## 依赖顺序

1. 任务 1 是所有现有文件修改的前置条件。
2. 任务 2 先恢复可信的全量测试基线。
3. 任务 3 提供任务 4–7 反复保存状态所需的可靠底座。
4. 任务 4 建状态模型，任务 5 建调度，任务 6 建文件闭环，必须依次完成。
5. 任务 7 接入主工作流后，任务 8 才能按实际行为更新整个 skill。
6. 任务 9 必须在自动测试、包校验、编码验证全部通过后执行。

## 主要风险与控制

- 注协接口可能不允许同一 `batch_id` 多次请求不同维度：先以第一批成功后对同一批次发第二批做实测；若接口明确拒绝，只把“上传企业名单”改为每个维度组重新上传，分组、成功跳过和关口规则不变，并先向用户报告这一硬限制。
- 下载任务缺少 `task_id` 或 `batch_id`：继续沿用触发前任务快照、触发时间和唯一新任务的保守关联，出现多个候选立即暂停。
- 空数据与漏文件混淆：只有存在对应工作簿且只有表头才记为明确无数据；文件不存在永远不算完成。
- 批次之间同名文件冲突：比较 SHA-256，同哈希复用，不同哈希暂停，绝不覆盖。
- Firefox 登录在长时间轮询中失效：保存全部进度，重新登录后从未完成维度恢复，不重下成功维度。
- 服务器节流或慢任务：人工节奏和两个完整窗口优先；只拆失败组，不并发轰炸接口。

## 第 3 阶段自检

- 需求覆盖：20 维、10+10、逐级拆分、成功不重下、两轮窗口、三类人工暂停、单维暂停、完成关口、旧任务保留、四川江擎实测均有对应任务。
- 无模糊占位：每个实现任务都给出具体文件、接口、失败测试、最小实现和验证命令。
- 接口一致：状态字段统一为 `required_dimensions`、`dimension_results`、`pending_groups`、`active_group`、`batch_history`、`poll_windows`；导出闭环统一由 `run_to_completion()` 驱动。
- 写入边界：只修改本计划列出的文件；不删除、不提交 Git、不启用子任务、不新增依赖。
