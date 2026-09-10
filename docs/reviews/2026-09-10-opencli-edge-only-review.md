# 验收审查:OpenCLI + Edge 唯一路线

日期: 2026-09-10
范围: 登录与取数路线重构(Firefox/Cookie 复制退役,opencli 页面取数通道上线)

## 审查方法

1. 实现逐项对照需求文档、设计文档、实施计划。
2. 真机端到端验收(auth 编排 + 官方接口取数)。
3. 全量自动化测试 + 分发包完整性校验。

## 审查发现与处置

| 发现 | 严重度 | 处置 |
|---|---|---|
| transport 模块 `json` 参数遮蔽标准库 | 高 | 模块别名 `jsonlib`,测试覆盖 |
| `.format()` 拼 JS 模板花括号转义脆弱 | 高 | 改为占位符替换 |
| 单测触达真机 opencli(CLI 定位未注入) | 中 | 测试统一注入假 cli_path |
| 页面不在知识库域时 expired 文案误导 | 低 | 文案改为"重新执行登录引导",引导会自动导航,无需重复输密码 |
| SKILL.md 否定句"不复制 Cookie"命中禁词子串"复制 Cookie" | 低 | 措辞改"不提取 Cookie" |
| 契约词"只保存到当前 Windows 用户"与新语义不符 | 中 | validate_bundle 与测试契约词更新为"不保存任何凭据" |

## 对照需求验收

1. ✅ 登录页只出现在 Edge:`opencli browser <session> open` 打开,与系统默认浏览器无关。
2. ✅ 无需复制 Cookie:真机实测 `auth_status` → authenticated,`get_user_index` → HTTP 200 / status_code==0。
3. ✅ 全量测试 141 通过 + 57 子测试;client.py 与 exporter.py 源码零改动(transport 鸭子类型兼容)。
4. ✅ `validate_bundle.py` BUNDLE_VALID=True;SKILL.md/README/user-flow 不再有"首选 Firefox"表述,OpenCLI 命名澄清已写入。
5. 交付:提交并推送 GitHub(用户已明确授权)。

## 真机验收证据(2026-09-10)

- `opencli doctor`:Daemon 19825 / Extension v1.0.24 / profile vzymxgmh 全部 OK。
- `collect_browser_session` → `auth_status()` = authenticated。
- `OpenCliTransport.request("GET", .../get_user_index)` → 200 / status_code==0。
- 旧路线同接口(复制可见 Cookie + Python 直连)恒为 401,已由实验 2 证实并记录于需求文档。

## 残余风险

1. 大体积 ZIP 经 base64 过 CLI 标准输出的上限未在真实 20 维导出中压测;既有失败恢复
   (连续三次失败转逐项下载、任务状态保留)可兜底,建议首次真实导出时留意。
2. opencli CLI 为新增一次性依赖;未安装机器的 preflight 文案已给出指引。
3. 旧 auth.bin 成为无害遗留文件,代码不再读写。
