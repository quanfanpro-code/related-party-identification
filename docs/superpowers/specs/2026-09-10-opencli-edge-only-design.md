# 设计:OpenCLI + Edge 唯一路线(页面上下文取数)

日期: 2026-09-10
需求: docs/requirements/2026-09-10-opencli-edge-only-requirement.md

## 方案比较

| 方案 | 说明 | 结论 |
|---|---|---|
| A. 逆向守护进程 WebSocket 协议,直发 cookies 指令 | 实测握手 HTTP 400;协议非公开、随版本漂移 | 弃 |
| B. opencli network 抓请求 Cookie 头 | 输出仅响应 shape,无请求头 | 弃 |
| C. **页面上下文取数**:所有请求经 `opencli eval` 在页面内 fetch,浏览器自动携带全量 Cookie | WorkBuddy 实战与本轮实验双重验证有效;零凭据落盘 | **选定** |
| D. 给扩展文件打端口补丁做双连 | 破坏用户 opencli CLI 正常使用;维护魔改二进制 | 弃 |
| E. CDP 直连 Edge | 需要用户带调试端口重启 Edge,违背不打扰原则 | 弃 |

## 推荐方案架构

```
related_party_workflow.py(编排/状态机,基本不动)
        │
CicpaClient(限速/重试/缓存,零改动)
        │  self.session.request(...)  ←鸭子类型
OpenCliTransport(新, scripts/cicpa/browser_transport.py)
        │  subprocess: opencli --profile <p> browser <s> eval <js>
opencli CLI v1.8.7 → 守护进程 19825 → Edge 扩展 v1.0.24
        │  Runtime.evaluate / chrome.debugger
zsk-cmis.cicpa.org.cn 页面上下文 fetch(credentials:'include')
        │  自动携带 HttpOnly 全量 Cookie
中注协官方接口
```

核心决策:**OpenCliTransport 实现 `.request(method, url, **kwargs)` 并返回
带 `.status_code / .headers / .content / .json()` 的响应对象,即鸭子类型兼容
requests.Session 契约**。`CicpaClient.__init__` 已接受任意 session 注入,
client.py 与 exporter.py 因此零改动;限速、临时错误重试、429 退避、
401/403 语义全部原样生效。

## OpenCliTransport 契约

- 定位 CLI:`shutil.which("opencli")`;找不到抛 `OpenCliBridgeError`,文案给安装指引。
- 配置定位:环境变量 `RPI_OPENCLI_PROFILE` 优先;否则解析 `opencli profile list`,
  恰有一个连接中的配置时自动采用,零个/多个时报错(多个时提示设置环境变量)。
- 会话名:常量 `OPENCLI_SESSION = "related-party"`(租约绑定知识库标签页)。
- `.request(method, url, *, params=None, json=None, files=None, headers=None, timeout=30)`:
  - URL 必须以 `https://zsk-cmis.cicpa.org.cn` 开头(取数只发生在此域),
    相对化为 path+query,页面不在该域时先导航回来。
  - params 并入查询串;json 序列化为 body 并置 Content-Type;
    files(仅 `{"file": (name, bytes, mime)}` 一种形态)编译为 JS FormData+Blob;
    自动从 `document.cookie` 注入 `X-XSRF-TOKEN` 头。
  - JS 统一返回 `JSON.stringify({s: status, h: {...响应头子集}, b: base64(body)})`,
    Python 侧解码为响应对象(headers 需支持 `.get("Retry-After")`,
    content 为 bytes,json() 走 `json.loads(content)`);429 的 Retry-After 头
    由 fetch 的 `response.headers.get` 原样透传。
- 登录编排:
  - `open_page(url)`:`opencli browser <s> open <url>`;`get_url()` 读当前地址。
  - `verify_session()`:eval fetch `get_user_index`,status_code==0 即 True。
  - `wait_for_login(timeout=900, poll=2.5)`:确认在知识库域 → 轮询 verify_session,
    超时抛 `LoginTimeout`。SSO 跳转登录由站点自身完成,技能只打开知识库地址。

## 登录流程变化(auth.py)

- `collect_cookies_with_browser` → `collect_browser_session(bridge_factory=..., timeout, sleep, clock)`:
  检查 CLI/配置 → `open_page(ZSK_URL)` → `wait_for_login` → 返回 None(会话即登录态)。
- `run_guided_login`:调用上述编排,成功写 `authenticated`;失败写 `failed` 并带
  分类文案(CLI 缺失/配置缺失/等待超时/验证失败)。状态机 JSON 不变。
- `auth_status`:改为通过 `verify_session()` 实时验证,返回
  authenticated / expired / failed(附指引);不再读取 auth.bin。
- 删除:read_firefox_cookies、_read_mozlz4、default_browser_profile_path 的 Firefox 分支、
  detect_browser 的 Firefox 候选与默认浏览器探测(engine 概念随之简化)、
  open_default_browser、CredentialStore、DpapiProtector、_client_from_cookies、
  default_cookie_verifier、edge_bridge.py、read_browser_cookies.ps1。
- detect_browser 保留 Edge/Chrome 安装探测(preflight 仍需告知浏览器是否安装;
  executable_path 供 opencli-setup 打开扩展管理页使用)。

## workflow.py 变化

- imports 对齐 auth.py 新表面;auth_status 参数加 bridge 注入点(测试替身)。
- preflight:浏览器检测保留;新增 opencli CLI/配置连接检查(经注入的 checker);
  文案全面去 Firefox 化;needs_opencli_extension 分支保留(装了 Edge 但扩展未连接)。
- login-start `--browser {edge,chrome}`(删 firefox);opencli-setup 不变。

## 错误处理

| 失败 | 处置 |
|---|---|
| CLI 不存在 / 配置未连接 | OpenCliBridgeError,文案含一次性安装指引(CLI+扩展) |
| eval 超时 / 进程非零退出 | 包装为 OSError 子类,进 client 临时错误重试(1s/2s) |
| 登录等待超时 | LoginTimeout,login-status 置 failed,可重新 login-start |
| 页面被导航离知识库域 | transport 请求前检查 get_url,自动导航回知识库首页再重发一次 |
| HTTP 401(会话中途失效) | client 抛 AuthenticationRequired,workflow 既有 waiting_user 状态承接 |

## 测试策略

- 新增 test_browser_transport.py:CLI 定位与配置解析(注入 runner)、fetch JS 编译
  (GET/POST JSON/multipart/params)、eval 输出解析与 base64 解码、响应对象契约、
  wait_for_login 轮询与超时、域外导航回跳;全部经假 runner/假 eval 注入,不联网。
- test_auth.py:删除 Firefox/Cookie 复制/DPAPI 用例;新增登录编排用例
  (成功/超时/CLI 缺失/状态写入无敏感字段);`--browser-name` choices 断言更新。
- test_workflow.py:preflight 文案与 opencli 检查分支用例更新;
  login-start 参数边界(firefox 现为非法值);文档一致性用例同步。
- 真机冒烟(验收阶段):opencli 打开知识库 → 页面验证 status_code==0 →
  transport 发起一次真实 get_user_index。

## 迁移与回滚

- 纯代码+文档变更,单分支单提交;回滚 = git revert。
- 旧 auth.bin 留存无害;已保存的 Firefox 使用习惯由文档新表述替代。
