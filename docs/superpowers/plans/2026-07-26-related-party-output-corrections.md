# 关联方识别输出纠错实施计划

> **执行说明：**未获用户授权使用子代理，因此按 `executing-plans.md` 在当前任务中串行执行。每项任务使用复选框跟踪，严格遵循“失败测试 → 最小实现 → 全量验证”。未经用户单独授权，不执行提交、推送、合并或删除。

**目标：**修正原始导出位置、成果命名、汇总表格式和公开客户供应商误判，同时保持登录、限速、穿透深度、旧任务状态和既有导出器调用兼容。

**实现方式：**在线工作流先生成安全公司名和唯一成果路径，再把带公司名的原始导出目录交给注协导出器直接使用。客户、供应商表中的交易对手只进入候选发现，不进入风险引擎；无效的关系标记规则和页签整体移除。所有改动复用现有标准库、`openpyxl`、状态文件和测试结构，不引入新依赖。

**技术基础：**Windows、Python 3.9 及以上、标准库、`openpyxl`、`unittest`。

## 全局约束

- 只修改 `D:\BaiduSyncdisk\workbuddy skills\related-party-identification`。
- 当前分支为 `main`，存在一处本次工作开始前已有的 `scripts\cicpa\exporter.py` 未提交修改：相对下载地址统一补齐注协域名。后续补丁必须保留该修改。
- 代码实施前，先把本计划列出的全部既有文件备份到 `C:\Users\27651\BackUp\related-party-identification_实际时间戳\`，保留相对目录，并逐一核对 SHA-256。
- 新建文档和修改后的中文文档使用带签名的 UTF-8；供程序读取的 JSON 保持无签名 UTF-8。Python 源文件保持现有 UTF-8 形式，不得产生问号乱码或 Unicode 替换字符。
- 不安装依赖，不删除旧临时目录，不改写历史报告，不改变登录、限速、随机延迟、候选上限和股权穿透层数。
- 在线导出成果结构固定为：
  - `<公司名称>_关联方核查报告.xlsx`
  - `<公司名称>_主动发现候选清单.xlsx`
  - `<公司名称>_注协原始导出\`
- 同名成果已存在时增加时间戳，不覆盖旧结果。
- 已有数据模式只读取用户选择的目录，不复制、不移动用户原始数据，也不创建注协客户端。
- 公开客户和供应商只能作为候选来源；不能单独生成风险命中、风险等级或披露差异。
- 代码注释、测试名称、技能说明和用户文档统一使用“七类有效核查证据”，不再使用其他近似说法。
- 全部任务由同一个执行者串行修改，避免同一文件并发写入。
- 文件后的行号是计划编写时的定位快照。前序任务发生增删后，以同一条目写明的函数名、类名和测试名作为权威定位，不按漂移后的数字猜位置。

## 修改前恢复准备

- [ ] 记录当前分支、`git status --short` 和 `git diff -- scripts/cicpa/exporter.py`，确认已有相对下载地址修复仍在。
- [ ] 创建 `C:\Users\27651\BackUp\related-party-identification_实际时间戳\`。
- [ ] 按原相对目录备份以下既有文件：
  - `scripts\cicpa\exporter.py`
  - `scripts\related_party_workflow.py`
  - `scripts\discovery.py`
  - `scripts\related_party_check.py`
  - `scripts\tests\test_exporter.py`
  - `scripts\tests\test_workflow.py`
  - `scripts\tests\test_discovery.py`
  - `scripts\tests\test_related_party_check.py`
  - `.gitignore`
  - `SKILL.md`
  - `README.md`
  - `NOTICE`
  - `references\SOURCES.json`
  - `references\rules.md`
  - `references\dimensions.md`
  - `references\cases.md`
  - `references\user-flow.md`
- [ ] 对源文件和备份文件分别计算 SHA-256；只有文件数量一致且每一组哈希相同后才能开始任务一。
- [ ] 把备份目录、哈希清单和修改前差异保存为本轮恢复证据。恢复时只从该备份目录复制本轮涉及文件，不运行重置、清理或永久删除命令。
- [x] 已把真实验收样例从系统临时目录复制到稳定位置：

```text
C:\Users\27651\BackUp\related-party-identification-plan_20260726_232811\acceptance-input\complete-dimensions_files
```

复制结果为源文件 17 个、备份文件 17 个、哈希不一致 0 个、合计 74,814 字节。任务六只使用这个稳定副本，不再依赖可能被系统清理的临时目录。

---

### 任务零：恢复修改前已经损坏的测试基线和分发格式

**批准来源：**修改前完整测试发现共同故障后，用户已经明确批准把该修复纳入本轮。

**文件：**

- 修改：`scripts` 下全部 22 个中文 Python 文件的编码
- 修改：`scripts\tests\test_auth.py`
- 修改：`scripts\tests\test_client.py`
- 修改：`scripts\tests\test_discovery.py`
- 修改：`scripts\tests\test_edge_bridge.py`
- 修改：`scripts\tests\test_exporter.py`
- 修改：`scripts\tests\test_opencli_setup.py`
- 修改：`scripts\tests\test_related_party_check.py`
- 修改：`scripts\tests\test_sync_upstream.py`
- 修改：`scripts\tests\test_workflow.py`
- 修改：`SKILL.md`
- 修改：`README.md`

**修改前失败证据：**

- 完整测试运行 81 项，出现 21 个错误和 1 个失败。
- 9 个测试文件的首条导入语句被错误写成 `#import` 或 `#from`。
- 技能包校验报告 22 个中文 Python 文件和 `README.md` 缺少要求的 UTF-8 签名。
- `SKILL.md` 的许可证注释放在 YAML 配置头之前，导致技能配置无法识别。

- [ ] 恢复 9 条被误注释的导入语句，只删除错误的注释井号。
- [ ] 把 `SKILL.md` 的 YAML 配置头移动到文件第一行，许可证说明紧随配置头之后，正文保持原顺序。
- [ ] 按 `validate_bundle.py` 的现有规则，把 22 个中文 Python 文件和 `README.md` 保存为带签名的 UTF-8；`SKILL.md` 保持无签名 UTF-8。
- [ ] 运行 `python -X utf8 -m unittest discover -s scripts/tests -p "test_*.py" -v`，确认导入错误消失。
- [ ] 运行 `python -X utf8 scripts/validate_bundle.py`，确认编码和技能配置头错误消失。
- [ ] 若仍有与本任务无关的原有失败，只调查共同根因，不顺手修改业务逻辑；无法收敛时再向用户报告。

---

### 任务一：把公开客户和供应商正确提取为候选

**文件：**

- 修改：`scripts\discovery.py:36-41,156-163`
- 修改：`scripts\related_party_workflow.py:424-484`
- 测试：`scripts\tests\test_discovery.py:113-153`
- 测试：`scripts\tests\test_workflow.py:306-410`

**接口：**

- 输入：`extract_seed_export_candidates(data_dir) -> List[SeedExportCandidate]`
- 输出：`SeedExportCandidate(name: str, relation_type: str, red_flags: tuple = ())`
- 关系类型固定为“公开客户关系”或“公开供应商关系”。
- `discover()` 继续把这些主体放入第一层候选，但不再接受或生成 `marked_related`。

- [ ] **步骤一：增加真实字段结构的失败测试**

在 `test_workflow.py` 增加一个客户表和一个供应商表。第二列“公司名称”写被审计单位，第八列“关联方名称”写公开交易对手：

```python
def make_seed_counterparty_export(
    path,
    owner_name,
    relation_label,
    counterparty_name,
):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append([
        "序号", "公司名称", "公告时间", "金额", "占比（%）",
        "与本公司关系", "货币代码", "关联方名称", "关联方ID", "关联理由",
    ])
    sheet.append([
        1, owner_name, "2026-07-01", 100, 10,
        relation_label, "CNY", counterparty_name, "org-1", "公开公告",
    ])
    workbook.save(path)
```

增加测试并明确第八列才是应提取的交易对手，第二列被审计单位不能被误作候选：

```python
def test_客户供应商第八列是公开交易对手而不是关联方认定(self):
    artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_counterparty_extract_"))
    make_seed_counterparty_export(
        artifact_dir / "客户.xlsx",
        "甲公司",
        "客户",
        "客户公司",
    )
    make_seed_counterparty_export(
        artifact_dir / "供应商.xlsx",
        "甲公司",
        "供应商",
        "供应商公司",
    )

    candidates = extract_seed_export_candidates(artifact_dir)

    self.assertEqual(
        [(item.name, item.relation_type) for item in candidates],
        [("客户公司", "公开客户关系"), ("供应商公司", "公开供应商关系")],
    )
    self.assertNotIn("甲公司", [item.name for item in candidates])
```

- [ ] **步骤二：复现注协工作簿范围信息错误**

再构造一份工作表内容有十列、但工作表 XML 的 `dimension` 错写成 `A1:A2` 的文件。使用 `zipfile` 生成新测试文件，不删除原文件：

```python
def copy_with_wrong_dimension(source, destination):
    with zipfile.ZipFile(source, "r") as input_zip:
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as output_zip:
            for item in input_zip.infolist():
                content = input_zip.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    text = content.decode("utf-8")
                    text = re.sub(
                        r'<dimension ref="[^"]+"',
                        '<dimension ref="A1:A2"',
                        text,
                        count=1,
                    )
                    content = text.encode("utf-8")
                output_zip.writestr(item, content)
```

测试仍应提取第八列交易对手。旧实现使用只读模式时只能看到 A 列，因此该测试必须失败。

- [ ] **步骤三：运行定向测试并确认失败原因**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_workflow.WorkflowModeTests.test_客户供应商第八列是公开交易对手而不是关联方认定 -v
python -X utf8 -m unittest scripts.tests.test_workflow.WorkflowModeTests.test_错误工作表范围信息不影响交易对手提取 -v
```

预期：旧实现提取不到“客户公司”“供应商公司”，或错误返回“甲公司”；失败原因必须与字段语义和只读范围一致。

- [ ] **步骤四：增加候选理由的失败测试**

把 `test_discovery.py` 中旧的 `marked_related=True` 测试改成：

```python
export_candidates = [
    SeedExportCandidate(name="客户公司", relation_type="公开客户关系"),
    SeedExportCandidate(name="供应商公司", relation_type="公开供应商关系"),
]
```

并验证：

```python
self.assertIn(
    "被审计单位完整维度导出：公开客户关系",
    result.candidates["客户公司"].reasons,
)
self.assertNotIn(
    "导出表内关联标注",
    result.candidates["客户公司"].reasons,
)
```

旧数据类和旧理由仍存在时，该测试必须失败。

- [ ] **步骤五：实施最小修复**

在 `SeedExportCandidate` 中删除 `marked_related`。`discover()` 直接使用 `relation_type` 形成候选理由。

在 `extract_seed_export_candidates()` 中：

1. 以普通模式打开工作簿，避开注协文件错误 `dimension` 对只读模式的截断。
2. 只把“关联方名称”列解释为该行公开交易对手名称；“公司名称”列仅表示公告归属公司。
3. 客户表生成“公开客户关系”，供应商表生成“公开供应商关系”。
4. 缺少交易对手名称的行跳过，不再回退为被审计单位名称。
5. 继续按交易对手名称去重。

最小核心逻辑为：

```python
counterparty_index = next(
    (
        index
        for index, header in enumerate(headers)
        if header == "关联方名称"
    ),
    7 if len(headers) >= 8 else None,
)
```

- [ ] **步骤六：运行任务一测试**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_workflow scripts.tests.test_discovery -v
```

预期：任务一新增测试和原有候选发现、两层穿透、候选上限测试全部通过。

---

### 任务二：彻底移除公开交易对手造成的硬关联误判

**文件：**

- 修改：`scripts\related_party_check.py:365-390,465-471,543,675-707,985-996,1129-1144`
- 测试：`scripts\tests\test_related_party_check.py:18-24,130-184`

**接口：**

- 保留：`run_check(..., disclosed_parties=None, ...)` 的参数，保证已有调用和旧任务兼容。
- 删除：`rule4_disclosed_gap()`、`Company.customers_marked_related`、`Company.suppliers_marked_related`。
- 删除报告页签：`04_注协标记关联`、`04b_用户披露差异`。
- 保留其他证据页签当前名称和编号，避免无关的下游兼容性变化。

- [ ] **步骤一：把旧错误语义测试替换为失败回归测试**

在 `test_related_party_check.py` 中新增两个仅供该测试模块使用的辅助函数。它们与任务一的 `make_seed_counterparty_export()` 相互独立，名称和签名不混用：

```python
def make_check_basic_export(directory, company_names):
    path = Path(directory) / "基础工商信息.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["字段{}".format(index) for index in range(1, 24)])
    for company_name in company_names:
        sheet.append(blank_basic_row(company_name))
    workbook.save(path)
    return path


def make_check_counterparty_export(
    path,
    owner_name,
    relation_label,
    counterparty_name,
):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append([
        "序号", "公司名称", "公告时间", "金额", "占比（%）",
        "与本公司关系", "货币代码", "关联方名称", "关联方ID", "关联理由",
    ])
    sheet.append([
        1, owner_name, "2026-07-01", 100, 10,
        relation_label, "CNY", counterparty_name, "org-1", "公开公告",
    ])
    workbook.save(path)
    return path
```

建立只包含“甲公司”基础工商信息的导出目录，再建立客户表，其中第八列为“乙公司”。分别在没有和提供用户自报名单时运行核查：

```python
def test_公开客户供应商不会生成关联风险或披露差异(self):
    artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_public_counterparty_"))
    make_check_basic_export(artifact_dir, ["甲公司"])
    make_check_counterparty_export(
        artifact_dir / "客户.xlsx",
        owner_name="甲公司",
        relation_label="客户",
        counterparty_name="乙公司",
    )

    result = run_check(
        data_dir=artifact_dir,
        target_names=["甲公司"],
        output_path=artifact_dir / "甲公司_关联方核查报告.xlsx",
        as_of_date=date(2026, 7, 1),
        disclosed_parties={"丙公司"},
    )

    self.assertEqual(result.hits, [])
    self.assertEqual(result.summary, [])
```

旧实现会生成“注协标记但未在用户自报名单出现”的硬关联，因此测试必须失败。

- [ ] **步骤二：增加报告页签失败断言**

在同一测试中打开报告并验证：

```python
self.assertNotIn("04_注协标记关联", workbook.sheetnames)
self.assertNotIn("04b_用户披露差异", workbook.sheetnames)
```

旧实现固定创建两个页签，因此断言必须失败。

- [ ] **步骤三：运行定向测试并确认失败**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_related_party_check.DisclosedPartySemanticsTests -v
```

预期：旧实现返回硬关联命中并创建两个错误页签。

- [ ] **步骤四：实施最小删除**

从 `Company`、`build_company()`、规则区、`run_check()` 聚合路径和 `dim_sheets` 中删除全部错误关系标记逻辑。保留 `disclosed_parties` 参数及“是否提供用户自报名单”的任务范围记录，只作为旧调用兼容信息，不参与自动风险判断。

同时把文件开头和规则区注释中的“八层规则”统一改为“七类有效核查证据”。不得改动股权、人员、工商指纹、客商画像、历史、担保资金链和无形资产规则。

- [ ] **步骤五：扫描所有错误调用点**

运行：

```text
rg -n --no-ignore-files "rule4_disclosed_gap|marked_related|customers_marked_related|suppliers_marked_related|注协标记关联关系|用户披露差异" scripts
```

预期：运行代码和测试中均无旧字段、旧函数和旧证据名称。

- [ ] **步骤六：运行核查引擎测试**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_related_party_check -v
```

预期：短行边界、动态成立月份、零参保、错误记录、公开交易对手不误判及页签测试全部通过。

---

### 任务三：修正汇总表头和第一条数据格式

**文件：**

- 修改：`scripts\related_party_check.py:934-984`
- 测试：`scripts\tests\test_related_party_check.py:185-215`

**接口：**

- 保持 `write_report()` 参数和工作表名称不变。
- 固定第三行为表头，第四行起为数据，冻结窗口为 `A4`。

- [ ] **步骤一：增加失败格式测试**

直接用一条汇总结果调用 `write_report()`：

```python
summary = [{
    "company_a": "乙公司",
    "company_b": "甲公司",
    "relation_type": "审计对象-交易对手",
    "is_related": "是（建议确认）",
    "max_level": "🔴硬关联",
    "dimensions": "股权控制穿透",
    "hit_count": 1,
    "evidence": "甲公司投资乙公司",
    "suggestion": "核对股权资料",
}]
```

加载生成文件并验证：

```python
self.assertEqual(sheet["A3"].value, "公司A")
self.assertEqual(sheet["A3"].fill.fgColor.rgb, "00305496")
self.assertEqual(sheet["A3"].font.color.rgb, "00FFFFFF")
self.assertEqual(sheet["A4"].value, "乙公司")
self.assertNotEqual(sheet["A4"].fill.fgColor.rgb, "00305496")
self.assertEqual(sheet.freeze_panes, "A4")
```

旧实现会把表头样式加到第四行并冻结到第五行，因此测试必须失败。

- [ ] **步骤二：运行测试并确认失败位置**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_related_party_check.RunCheckIntegrationTests.test_汇总第三行是表头且第一条数据使用风险格式 -v
```

预期：第三行没有蓝底，第四行误用表头样式，冻结位置为 `A5`。

- [ ] **步骤三：修正行号**

保持 `start_row = 3`，把表头样式应用到 `start_row`；数据着色从 `start_row + 1` 开始；行高应用到第三行；冻结位置改为 `A4`：

```python
for c in range(1, len(headers) + 1):
    cell = ws0.cell(row=start_row, column=c)
    cell.fill = FILL_HEADER
    cell.font = FONT_HEADER
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )
ws0.row_dimensions[start_row].height = 28

for ri, item in enumerate(summary):
    row_idx = start_row + 1 + ri

ws0.freeze_panes = "A4"
```

- [ ] **步骤四：运行格式和核查测试**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_related_party_check -v
```

预期：表头、数据行、冻结位置及任务二的误判回归全部通过。

---

### 任务四：生成带公司名的成果，并让注协直接导出到成果目录

**文件：**

- 修改：`scripts\cicpa\exporter.py:74-105,378-416`
- 修改：`scripts\related_party_workflow.py:503-529,562-690`
- 测试：`scripts\tests\test_exporter.py:72-173`
- 测试：`scripts\tests\test_workflow.py:411-552`

**接口：**

- `CicpaExporter(client, *, staging_root=None, artifact_dir=None, now=...)`
- `_safe_company_stem(name: str) -> str`
- `run_workflow(..., exporter_factory: Optional[Callable] = None, ...)`
- 用户提供 `exporter_factory` 时仍按原签名 `exporter_factory(client)` 调用。
- `WorkflowResult.report_paths` 在线模式包含原始导出目录、候选清单和核查报告；已有数据模式不把用户输入目录当成输出。

- [ ] **步骤一：增加导出器成果目录失败测试**

把完整流程测试改为使用：

```python
delivery_dir = artifact_dir / "甲公司_注协原始导出"
exporter = CicpaExporter(
    client,
    artifact_dir=delivery_dir,
    now=lambda: "2026-07-26 15:00",
)
```

并验证：

```python
self.assertEqual(Path(state.staging_dir), delivery_dir)
self.assertEqual(extract_dir, delivery_dir)
self.assertTrue((delivery_dir / "基础工商信息.xlsx").exists())
self.assertTrue((delivery_dir / "complete-dimensions.zip").exists())
self.assertTrue((delivery_dir / "upload-companies.xlsx").exists())
self.assertFalse((delivery_dir / "complete-dimensions_files").exists())
```

旧构造器不接受 `artifact_dir`，测试必须失败。

- [ ] **步骤二：增加工作流命名失败测试**

已有数据模式验证报告名：

```python
self.assertEqual(
    result.report_paths,
    [artifact_dir / "甲公司_关联方核查报告.xlsx"],
)
```

主动发现模式验证：

```python
self.assertTrue(
    (artifact_dir / "甲公司_主动发现候选清单.xlsx").exists()
)
self.assertIn(
    artifact_dir / "甲公司_关联方核查报告.xlsx",
    result.report_paths,
)
```

增加名称清理边界：

```python
self.assertEqual(_safe_company_stem(" 甲<乙>:公司. "), "甲_乙__公司")
self.assertEqual(_safe_company_stem("CON"), "_CON")
self.assertEqual(_safe_company_stem("..."), "未命名公司")
```

再预先创建同名报告，验证新结果增加时间戳且旧文件内容不变。

- [ ] **步骤三：验证默认导出器收到命名目录**

使用 `unittest.mock.patch` 替换工作流模块中的 `CicpaExporter`，调用未传 `exporter_factory` 的在线工作流，并验证：

```python
constructor.assert_called_once_with(
    client,
    artifact_dir=artifact_dir / "甲公司_注协原始导出",
)
```

同时验证用户传入 `lambda client: fake_exporter` 时仍只接收一个 `client` 参数，现有测试替身不需要改签名。

- [ ] **步骤四：运行任务四失败测试**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_exporter.ExportFlowTests.test_真实流程保存批次号并只下载对应任务 -v
python -X utf8 -m unittest scripts.tests.test_workflow.WorkflowModeTests -v
```

预期：旧构造器拒绝 `artifact_dir`，旧成果名不含公司名，默认工作流未传入命名目录。

- [ ] **步骤五：实现安全公司名**

在工作流中使用标准库 `re`，不增加依赖：

```python
_WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *{"COM{}".format(index) for index in range(1, 10)},
    *{"LPT{}".format(index) for index in range(1, 10)},
}


def _safe_company_stem(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name).strip())
    cleaned = cleaned.rstrip(" .")[:80] or "未命名公司"
    if cleaned.upper() in _WINDOWS_RESERVED_NAMES:
        cleaned = "_" + cleaned
    return cleaned
```

三个成果名称只从该函数生成一次。继续复用 `_safe_output_path()` 处理同名碰撞。

- [ ] **步骤六：实现导出器直达成果目录**

给 `CicpaExporter` 增加可选 `artifact_dir`。指定时 `_new_staging_dir()` 创建并返回该精确目录；未指定时保留现有 `staging_root + tempfile.mkdtemp()` 行为。

下载时只有当状态中的 `staging_dir` 与本次 `artifact_dir` 相同时，才直接解压到成果目录；恢复旧任务的临时状态仍使用原来的 `complete-dimensions_files` 子目录：

```python
direct_delivery = (
    self.artifact_dir is not None
    and staging_dir.resolve() == self.artifact_dir.resolve()
)
extract_dir = (
    staging_dir
    if direct_delivery
    else staging_dir / "complete-dimensions_files"
)
```

保留现有相对下载地址修复：

```python
if not url.startswith("http"):
    url = self.ZSK_BASE + "/" + url.lstrip("/")
```

- [ ] **步骤七：接入在线工作流**

把 `exporter_factory` 默认值改为 `None`：

```python
if exporter_factory is None:
    exporter = CicpaExporter(client, artifact_dir=raw_export_dir)
else:
    exporter = exporter_factory(client)
```

在线模式把 `data_dir` 原始导出目录加入 `state.report_paths`；主动发现候选清单使用追加方式，不能覆盖已经记录的原始导出目录。已有数据模式只记录新生成报告。

- [ ] **步骤八：运行导出和工作流测试**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_exporter scripts.tests.test_workflow -v
```

预期：直接导出目录、旧暂存目录兼容、断点续跑、三个工作流入口、公司名清理、同名防覆盖和成果路径测试全部通过。

---

### 任务五：同步技能说明、字段映射、案例和分发说明

**文件：**

- 修改：`SKILL.md`
- 修改：`README.md`
- 修改：`NOTICE`
- 修改：`references\SOURCES.json`
- 修改：`references\rules.md`
- 修改：`references\dimensions.md`
- 修改：`references\cases.md`
- 修改：`references\user-flow.md`
- 修改：`.gitignore`
- 测试：`scripts\tests\test_workflow.py:553-674`

**接口：**

- 对话和分发文档统一使用 `CONTEXT.md` 与 `docs\adr\0001-公开客户供应商仅作候选.md` 中的术语。
- 当前证据框架称为“七类有效核查证据”。
- 三个成果名称使用带公司名模板。

- [ ] **步骤一：先修改文档契约测试并确认失败**

把旧的“注协标记关系与用户披露差异”契约替换为：

```python
required_phrases = [
    "候选不等于关联方",
    "红旗不等于硬关联",
    "公开客户关系",
    "公开供应商关系",
    "只作为候选",
    "不能单独形成风险证据",
    "初始 2 层",
    "最多 5 层",
    "只能从完整维度导出",
]
```

对 `SKILL.md`、`README.md`、`references\rules.md`、`references\dimensions.md` 和 `references\user-flow.md` 增加禁止断言：

```python
for phrase in [
    "注协标记关联关系",
    "导出表内关联标注",
]:
    self.assertNotIn(phrase, combined)
```

增加成果命名契约：

```python
for phrase in [
    "<公司名称>_关联方核查报告.xlsx",
    "<公司名称>_主动发现候选清单.xlsx",
    "<公司名称>_注协原始导出",
]:
    self.assertIn(phrase, combined)
```

运行：

```text
python -X utf8 -m unittest scripts.tests.test_workflow.SkillDocumentationContractTests -v
```

预期：旧文档仍包含错误规则和旧固定文件名，测试失败。

- [ ] **步骤二：局部修正技能入口和用户说明**

`SKILL.md`：

- 把“八个维度”改为“七类有效核查证据”。
- 把参考区中“八层规则详解”的旧名称同步改为“七类有效核查证据详解”，不能遗漏当前文件第 135 行附近的引用。
- 明确公开客户、供应商只扩充候选，不能单独形成风险。
- 输出部分列出三个带公司名的成果模板。
- 删除“注协标记关系”“用户披露差异”的执行指令。
- 把“亲属关系（第9维度）”改为不带错误编号的公开数据限制说明。

`README.md`：

- 第八步删除“注协数据是否标记了关系”，改为“公开客户、供应商只作为候选，继续核对独立证据”。
- “最终会拿到什么”列出三个实际名称模板及在线、离线边界。
- 说明原始导出表与报告在同一用户选择目录。
- 保留登录、Firefox、OpenCLI、限速和安装说明原文，不扩大本次范围。

`references\user-flow.md`：

- 结果交付话术使用带公司名成果。
- 删除旧的关系标记和披露差异话术。
- 增加“公开客户、供应商已列入候选，但不会因此直接判定关联方”的解释。

- [ ] **步骤三：局部修正规则、字段和案例**

`references\rules.md`：

- 标题改为“七类有效核查证据详解”。
- 整段删除错误的维度四关系标记规则。
- 增加“公开交易对手候选来源”边界说明。
- 其余规则内容不改变阈值；汇总表只列七类有效核查证据。

`references\dimensions.md`：

- 客户、供应商表 H 列解释为“公开交易关系另一方名称”。
- 用途改为主动发现候选，不映射风险规则。
- 删除规则四映射行。

`references\cases.md`：

- 保留监管案例中已经由外部事实认定的关联方结论。
- 删除“客户表 H 列本身证明关联方”和“当前规则覆盖注协标记关系”的错误推导。
- 改写为“公开客户身份只是调查入口，仍需股权、人员、联系方式、资金等独立证据”。

`NOTICE` 和 `references\SOURCES.json`：

- 把本地改造说明中的“关系标记与用户披露差异边界”改为“公开客户供应商候选与风险证据边界”。
- 不改变上游仓库、提交号和许可证事实。

- [ ] **步骤四：先用失败测试锁定 `.gitignore` 坏规则**

在 `SkillDocumentationContractTests` 增加：

```python
def test_gitignore_不含搜索工具无法解析的末尾反斜杠(self):
    lines = (self.skill_root / ".gitignore").read_text(
        encoding="utf-8-sig"
    ).splitlines()
    invalid = [
        line
        for line in lines
        if line.strip() and not line.lstrip().startswith("#") and line.endswith("\\")
    ]
    self.assertEqual(invalid, [])
```

先运行并确认当前第 17 行 `%LOCALAPPDATA%\related-party-identification\` 使测试失败。随后只删除这一条无效规则；Git 忽略文件不会展开 Windows 环境变量，而且目标运行数据本来就在仓库外，因此不需要替代规则。

- [ ] **步骤五：验证文档契约、搜索工具与编码**

运行：

```text
python -X utf8 -m unittest scripts.tests.test_workflow.SkillDocumentationContractTests -v
python -X utf8 scripts/validate_bundle.py
rg -n "公开客户关系|公开供应商关系" .
```

预期：文档契约通过；技能包输出 `BUNDLE_VALID=True`；普通 `rg` 搜索正常完成，不再出现 `.gitignore` 解析警告。

检查全部本轮中文文档：

- UTF-8 解码成功。
- Markdown 和文本文件以 UTF-8 签名开头。
- `references\SOURCES.json` 不带签名且可由 `json.loads()` 读取。
- 不含 Unicode 替换字符。

---

### 任务六：完整回归和真实样例验收

**文件：**

- 验证：本计划全部修改文件
- 新建：`docs\reviews\2026-07-26-related-party-output-corrections-review.md`，在第五阶段按验收结果填写

**接口：**

- 使用全部自动测试验证行为。
- 使用用户刚才真实运行留下的注协导出数据验证误判消失。
- 不联网、不重新提交注协导出任务，不覆盖用户现有报告。

- [ ] **步骤一：运行完整测试**

在技能根目录运行：

```text
python -X utf8 -m unittest discover -s scripts/tests -p "test_*.py" -v
```

预期：全部测试通过，退出码为 0；记录实际测试数量。

- [ ] **步骤二：运行技能包和工作区检查**

运行：

```text
python -X utf8 scripts/validate_bundle.py
git diff --check
git status --short
```

预期：

- `BUNDLE_VALID=True`
- `git diff --check` 退出码为 0
- 变更范围只包含计划列出的文件和本轮新增需求、设计、决策、计划、审查文档
- `scripts\cicpa\exporter.py` 原有相对下载地址修复仍然存在

- [ ] **步骤三：用真实注协导出数据生成新报告**

使用阶段三已经复制并逐文件核对哈希的稳定目录：

```text
C:\Users\27651\BackUp\related-party-identification-plan_20260726_232811\acceptance-input\complete-dimensions_files
```

被审计单位：

```text
泸州骐骊智能系统技术有限公司
```

在新的验收临时目录中生成：

```text
泸州骐骊智能系统技术有限公司_关联方核查报告.xlsx
```

不写入用户原成果目录，不覆盖旧报告。

- [ ] **步骤四：核对真实样例结果**

程序化检查：

1. 风险命中中不存在“注协标记关联关系”或“用户披露差异”。
2. 四川融创嘉科技有限公司、宁波未有智行科技有限公司、江苏中科重德智能科技有限公司等公开交易对手，不再仅凭客户供应商表进入硬关联。
3. 汇总页第三行是蓝底白字表头，冻结位置为 `A4`。
4. 若存在由其他独立规则形成的数据行，从第四行开始使用各自风险颜色。
5. 报告可由 `openpyxl` 重新打开，主要工作表和标题完整。

- [ ] **步骤五：复核需求、设计和实施计划**

逐项对照：

- `docs\requirements\2026-07-26-related-party-output-corrections-requirement.md`
- `docs\superpowers\specs\2026-07-26-related-party-output-corrections-design.md`
- 本实施计划
- `CONTEXT.md`
- `docs\adr\0001-公开客户供应商仅作候选.md`

把发现、修复、测试命令、退出码、测试数量、真实样例结果和残余限制写入审查文档。发现问题时回到对应任务继续按失败测试、最小修复、重新验证处理。

- [ ] **步骤六：交付前新鲜验证**

完成审查修复后，再运行一次：

```text
python -X utf8 -m unittest discover -s scripts/tests -p "test_*.py" -v
python -X utf8 scripts/validate_bundle.py
git diff --check
```

只有三项都刚刚成功，且真实样例确认四个问题已经消失，才能进入交付阶段。交付时列出备份目录、实际变更文件、测试数量、真实样例路径、未解决限制和恢复方法。

## 依赖顺序

1. 修改前恢复准备。
2. 任务零先恢复可执行、可验证的修改前基线。
3. 任务一纠正候选提取语义，为后续文档和端到端测试提供正确输入。
4. 任务二移除共同误判根因。
5. 任务三修正报告表现层。
6. 任务四修正成果目录和命名。
7. 任务五统一用户说明和领域语义。
8. 任务六做完整回归、真实样例验收和第五阶段审查。

任务一与任务二都涉及公开客户供应商语义，必须按顺序执行；任务二和任务三虽然修改同一个核查文件，也必须串行完成。任务四依赖命名约定和已确认的直接导出方案。文档只能在运行行为稳定后统一更新。

## 主要风险和控制

- **已有未提交修改被覆盖：**修改前保存差异和哈希，后续只用局部补丁，最终重新核对相对下载地址修复。
- **注协导出工作簿范围信息错误：**用真实字段布局和错误 `dimension` 的测试覆盖，读取方式按实际文件校准。
- **旧任务无法恢复：**`artifact_dir` 是可选项；状态中的旧临时目录继续按旧布局解压。
- **离线模式复制用户数据：**只在在线默认导出器创建命名目录；已有数据模式路径保持只读输入。
- **客户供应商仍以其他规则命中：**允许独立证据产生风险，但报告证据中不得再出现关系标记或披露差异。
- **同名成果覆盖：**复用当前时间戳碰撞处理，并验证旧文件内容不变。
- **文档再次诱导误判：**用自动文档契约同时检查必备新表述和禁止旧表述。
- **搜索工具被坏忽略规则阻断：**用失败测试锁定末尾反斜杠规则，删除无效行后用普通 `rg` 实际验证。
- **用户当前使用的是链接目录：**在当前技能仓库直接局部修改，修改会立即反映到 WorkBuddy 已登记的技能路径；不复制出另一份运行副本。

## 计划自检

- 需求四个问题均有独立失败测试、最小实现和验收步骤。
- 在线、离线、断点恢复、同名碰撞和 Windows 非法文件名均有明确边界。
- 错误规则的代码、数据字段、测试、页签、对话话术、规则说明、字段映射、案例说明和来源说明均已覆盖。
- 任务一和任务二使用的三个测试辅助函数均有完整定义，名称和参数不再冲突。
- 真实验收样例已有稳定副本和逐文件哈希证据，不依赖系统临时目录。
- `.gitignore` 第 17 行错误已纳入带失败测试的最小修复。
- 没有新依赖、并行代理、自动提交、删除、清理、登录或联网动作。
- 所有函数名、参数名、成果模板和测试命令在前后任务中保持一致。
- 未留下待补内容或模糊的“稍后处理”步骤。
