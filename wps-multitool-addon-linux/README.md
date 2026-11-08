# wps-multitool-addon-linux

Linux/macOS 加载项通道的加载项侧（B2 骨架）。对应 Node 侧服务：`src/client/addon-channel.ts`（计划 §4.4/§12.1）。

## 结构（三型三目录，共享一份通道客户端）

```
addon-wps/   wps 型（文字）   manifest.xml + index.html + ribbon.xml + js/boot.js
addon-et/    et  型（表格）   同构
addon-wpp/   wpp 型（演示）   同构
js/channel-client.js          三型共用：握手 + 长轮询循环 + 结果回传 + 动作分发器
```

为什么必须三个：S1 实测三型对象模型互相看不见（wps 只有 Documents / et 只有 Workbooks /
wpp 只有 Presentations），一个加载项只能拿一种应用的对象模型（计划 §12.1）。
三型连**同一个端口**（默认 127.0.0.1:27740，`WPS_ADDON_PORT` 可改），靠握手 body 的
`type` 字段区分会话；同型后来者覆盖前者。

## 部署（Linux 真机，B3 联调时执行）

1. 三个目录各自拷入 `~/.local/share/Kingsoft/wps/jsaddons/<name>_<ver>/`
2. 改 `jsplugins.xml` 登记三个加载项（S1 探针已验证此路径可行）
3. 确认 `/opt/kingsoft/wps-office/office6/cfgs/oem.ini` 的 `[support] JsApiPlugin=true`
4. Node 侧起通道（`WPS_ADDON_TOKEN` 未设置时启动日志会打随机令牌）
5. WPS 里「工具 → WPS 加载项」**手动逐个打开三个加载项**（Linux 无自动加载，计划 §12.2），
   首次各点一次信任对话框（§12.3）
6. 服务端地址/令牌从 URL query 传给加载项页（`?server=http://127.0.0.1:27740&token=…`）

## 动作分发器（B2 骨架约定）

`js/channel-client.js` 的 `dispatch()`：ping/wireCheck 真实现，另有 word 5 + ppt 15 共 20 个动作已实现（S0/S1 验证过的集合，见 B2.5）；其余动作显式 unsupported，
**其余动作一律回 `{ ok:false, unsupported:'B3' }`** —— B3 才填 20+ 动作的 JSAPI 实现
（S1 覆盖率：支持 199 / 部分支持 45，适配器与实现顺序按 B4 结论）。
显式拒绝、绝不静默成功（计划 §6 D3 的精神）。

## 与 Windows 加载项（wps-multitool-addon/）的关系

骨架照它的 manifest/ribbon 结构改型别而来；**两者物理分开**（C2）：
Windows 走 COM 桥 + 27731 只读面板，Linux 走 27740 执行通道。
不要在两边互相搬实现。


## 双模型架构（B5-T4 落仓，2026-11-08）

12.x 的 webview 里 `window.Application` 是 WPS 注入的「同步 JSAPI 门面」（属性/方法直调，
返回值即时可得）；11.x 只有 jsAsynCall(cefQuery) 异步桥。两套模型共存于本包，boot 时选路：

```
boot() → WpsAddonModelDetect.apply(type)
  ├─ 探测到门面（12.x）→ WpsAddonActions{Word|Excel|Ppt} 替换为 js/actions-sync/ 同步实现
  └─ 探测不到（11.x） → 保留 js/actions/*.js 的 jsAsynCall 异步实现（兜底，基线不动）
```

- `js/actions-sync/runtime.js`：同步运行时（取集合用 .Item()、写值 Value2=、读值 Value()、
  裸  规整、AddEffect 禁读数量等 S1 §4 语义适配全部内建）
- `js/actions-sync/{word,ppt,excel}.js`：声明式动作规格（13/15/95 条），来自 S1-REPORT §3
- `js/model-detect.js`：双模型检测 + `POST /env` 上报（phase=modelDetect）
- 字节指纹：五文件 SHA-256 钉死在主仓 `scripts/manual/actions-sync-verify.js`（FAST 项）——
  与云端 B5-T4 真机验证版逐字节等价是 123/123 证据的前提，改动即门禁红
