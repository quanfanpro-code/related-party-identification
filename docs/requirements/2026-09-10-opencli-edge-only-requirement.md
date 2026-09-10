# 需求:登录与取数统一为 OpenCLI + Edge 唯一路线

日期: 2026-09-10
状态: 已确认(用户明确授权全程自主决策)
触发: WorkBuddy 实战对话失败复盘(见主对话)

## 背景与问题

2026-09-10 在 WorkBuddy 中执行"边际领域(大连/北京)主动发现"任务时,登录与取数连续失败。
经对本机环境与技能源码的完整排查和四组实测,确认三件事:

1. **登录页弹错浏览器(缺陷一)**。`collect_cookies_with_browser` 用 `open_default_browser`
   (即 `os.startfile`)打开注协登录页,页面总是弹在 Windows 系统默认浏览器(Firefox)里;
   `login-start --browser edge` 参数管不到"打开页面"这一步。用户被迫反复面对 Firefox 弹窗。
2. **Cookie 复制通道在本机不可用(缺陷二)**。技能的 Edge/Chrome 路线由内置桥
   (edge_bridge.py,自建 127.0.0.1:19825 WebSocket 服务)等扩展来送 Cookie;
   而本机装有 opencli CLI v1.8.7,其守护进程常驻监听同一端口 19825,扩展连接被守护进程持有,
   技能桥读到的登录令牌是旧值,官方接口一律 401。
3. **即使读到完整可见 Cookie,Python 直连也过不了验证(实测新发现)**。服务器校验
   HttpOnly Cookie,`document.cookie` 与扩展以外的任何通道都拿不齐;唯一被服务器认可的
   请求环境是**登录浏览器本身的页面上下文**。

同时,用户在本机使用 OpenCLI(opencli CLI + Edge 扩展)已有既定偏好:
"以后所有浏览器操作统一用 OpenCLI 驱动 Edge,不管 Edge 是不是系统默认浏览器"。

## 用户与使用场景

用户是注册会计师(非程序员),在 Windows 上通过 AI 对话驱动本技能完成关联方核查。
用户已在本机安装并常驻:opencli CLI v1.8.7 + 守护进程(端口 19825) + Edge 扩展 v1.0.24
(浏览器桥配置名 vzymxgmh)。Edge 始终登录中注协系统。

## 功能范围

1. 登录路线收敛:退役 Firefox 路线与 Cookie 复制路线,登录页打开、登录验证、取数请求
   全部经由 opencli CLI 驱动 Edge(扩展桥)完成。
2. 登录页打开改为 `opencli browser <session> open <URL>`,与系统默认浏览器无关。
3. 登录验证改为在页面上下文调用官方轻量接口(`get_user_index`,status_code==0 即通过)轮询。
4. 取数传输层改为"页面上下文 fetch":新增 OpenCliTransport,鸭子类型兼容
   requests.Session 的 `.request()` 契约,CicpaClient/CicpaExporter 零改动接入;
   JSON、multipart 上传(企业名单 xlsx)、ZIP 二进制下载均经同一通道。
5. 限速、重试、429 退避、任务状态机、断点续跑等既有安全机制全部保留(位于 client.py 编排层)。
6. 凭据不再落盘:CredentialStore/DPAPI/auth.bin 退役,浏览器会话即登录态。
7. 文件退役:edge_bridge.py、read_browser_cookies.ps1 及对应测试。
8. 文档同步:SKILL.md、README.md、references/user-flow.md、CONTEXT.md、NOTICE;
   澄清"技能所指 OpenCLI = opencli CLI 工具 + 其 Edge 扩展"这一命名二义性。

## 明确不做的内容

- 不逆向 opencli 守护进程的 WebSocket 协议(实验证明握手被拒,且属于逆向依赖,版本升级即碎)。
- 不魔改扩展文件做双端口双连(会破坏用户 opencli CLI 的正常使用)。
- 不引入 Playwright、不另建浏览器空间、不要求用户关闭已打开的浏览器。
- 不改动 20 维度导出、候选扩展、核查引擎、报告输出的任何业务逻辑。
- 不改 RateLimiter 的保守限速参数。

## 约束与风险

- ZIP 经 base64 走 CLI 标准输出,单批体积需在实测中确认上限(工商维度 ZIP 通常为
  数百 KB 级;若超过 CLI 限制,轮询分批或降级为单维度导出,既有失败恢复机制可覆盖)。
- 每次 API 调用伴随一次 opencli 进程启动(数百毫秒);技能限速为每分钟 15 次、
  普通查询随机等待 2-5 秒,进程开销相对等待可忽略。
- opencli CLI 未安装的机器:preflight 必须给出可执行的安装指引(安装 CLI + 扩展为一次性动作)。
- 已保存的旧 auth.bin 成为无害遗留文件,代码不再读写,文档注明即可。

## 验收标准

1. `login-start --browser edge` 后,登录页只出现在 Edge 中;系统默认浏览器为 Firefox 也不弹 Firefox。
2. 登录完成后技能无需复制任何 Cookie,即可通过官方接口完成取数验证
   (实测:页面上下文 get_user_index 返回 status_code==0)。
3. 全量测试通过;原 Firefox/Cookie 复制相关测试删除或替换为新路线测试;
   client.py 与 exporter.py 源码零改动。
4. `validate_bundle.py` 通过;SKILL.md/README/user-flow 描述与新路线一致,不再出现
   "首选 Firefox"表述,并写明 OpenCLI 命名澄清。
5. 变更提交至 GitHub 主分支。

## 实测证据摘要(2026-09-10)

| 实验 | 结果 | 结论 |
|---|---|---|
| 1. 页面上下文 fetch get_user_index | 200,status_code==0 | 浏览器会话有效,服务器认可页面请求 |
| 2. Python 携带 document.cookie 三件套直连(对齐 UA) | 401"未登陆用户" | 服务器校验 HttpOnly Cookie,复制可见 Cookie 不足 |
| 3. WebSocket 直连守护进程 19825 发 cookies 指令 | HTTP 400 握手拒绝 | 逆向守护进程协议不可行 |
| 4. opencli network 抓请求头 | 仅响应 shape,无请求头 | CLI 无读取 HttpOnly Cookie 的官方出口 |
| 附:opencli doctor / profile list / open / eval | 全部正常 | opencli CLI + Edge 扩展通道完整可用 |

## 自主决策记录(grilling 设计树自答)

- Q: Firefox 路线是否保留为备选? → 否,用户两次明确"只用 opencli 调 Edge 的路径,
  不要 Firefox 路线"。Chrome 走同一 opencli 通道,行为一致,代码保留其候选地位。
- Q: 取数传输放在哪一层? → transport 兼容 Session 接口,client/exporter 零改动
  (复用现有代码优先于新抽象,限速重试天然保留)。
- Q: 登录态是否仍落盘保存? → 否。页面即登录态,轮询验证即"每次调用先实际验证"
  的既有原则,且消除凭据落盘面。
- Q: edge_bridge 端口冲突是否要检测报错? → 模块整体退役,冲突不复存在。
