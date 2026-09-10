# 实施计划:OpenCLI + Edge 唯一路线

日期: 2026-09-10
设计: docs/superpowers/specs/2026-09-10-opencli-edge-only-design.md

串行执行(TDD:失败测试 → 最小实现 → 全量验证)。所有命令在技能根目录
`C:\Users\27651\Documents\BaiduSyncdisk\workbuddy skills\related-party-identification`
下执行,解释器用
`C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe -m pytest scripts/tests -q`。

## T1 browser_transport 新模块

- 文件: `scripts/cicpa/browser_transport.py`(新)、`scripts/tests/test_browser_transport.py`(新)
- 测试先行(全部注入假 runner,不联网):
  1. `_locate_cli` 找不到 opencli 时报错且文案含安装指引
  2. profile 解析:环境变量优先 / profile list 单配置自动采用 / 零个报错 / 多个报错提示环境变量
  3. fetch JS 编译:GET+params、POST json、multipart files、X-XSRF-TOKEN 注入、域校验
  4. eval 输出解析:裸文本与 JSON 引号包装两种形态;base64 解码为 bytes
  5. 响应对象契约:status_code / headers.get("Retry-After") / content / json()
  6. wait_for_login:先导航回域内再轮询;成功/超时两分支
- 成功判据:新模块测试全绿,不影响既有测试。

## T2 auth.py 重构

- 文件: `scripts/cicpa/auth.py`、`scripts/tests/test_auth.py`
- 测试先行:collect_browser_session 成功/超时/CLI 缺失;run_guided_login 状态写入
  (authenticated/failed,状态文件无敏感字段);`--browser-name` 只接受 edge/chrome
  对应的浏览器名;auth_status 经注入 bridge 验证返回三态。
- 实现:删 Firefox/Cookie 复制/DPAPI/CredentialStore/_client_from_cookies/
  default_cookie_verifier/open_default_browser;新增 collect_browser_session、
  新 auth_status(bridge 注入);_main choices 收窄。
- 删除文件: `scripts/cicpa/edge_bridge.py`、`scripts/cicpa/read_browser_cookies.ps1`、
  `scripts/tests/test_edge_bridge.py`(git rm)。
- 成功判据:auth 相关测试全绿;`grep -ri firefox scripts/` 无 Firefox 路线残留
  (文案性提及除外,应仅剩"不再使用 Firefox"类说明)。

## T3 workflow.py 对齐

- 文件: `scripts/related_party_workflow.py`、`scripts/tests/test_workflow.py`
- 测试先行:preflight 无浏览器/无扩展/未登录三分支新文案;login-start 传 firefox 报参数错误;
  auth-status 输出对接;既有"预检结果不包含认证秘密"保持通过。
- 实现:imports 对齐;preflight 加 opencli 检查注入点;login-start choices 收窄;
  run/resume 的 ensure_login 不变。
- 成功判据:workflow 测试全绿。

## T4 文档同步

- 文件: `SKILL.md`、`references/user-flow.md`、`CONTEXT.md`、`NOTICE`、
  `references/SOURCES.json`(如边缘说明涉及)
- 内容:登录状态节改写(OpenCLI + Edge 唯一路线、页面即登录态、不落盘凭据)、
  OpenCLI 命名澄清、失败恢复表更新(登录类)、反例红线去 Firefox。
- test_workflow 的文档一致性用例("对话手册覆盖三个入口""用户文案不暴露技术操作")
  驱动,失败→修改→通过。

## T5 验收

1. 全量测试 + validate_bundle.py。
2. 真机冒烟:opencli 打开知识库 → wait_for_login 通过(会话有效) →
   OpenCliTransport 真实 get_user_index 返回 status_code==0。
3. 写 `docs/reviews/2026-09-10-opencli-edge-only-review.md`。

## T6 交付

- README.md 增补"登录与取数路线(2026-09)"一节并修正旧表述。
- git add 全部变更;两个提交:①上游 SOURCES 基线更新(先前遗留),
  ②本次路线重构;push origin 主分支。
