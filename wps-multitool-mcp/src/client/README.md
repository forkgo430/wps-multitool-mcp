addon-routes.ts, addon-channel.ts, action-routes.ts, action-routes.generated.ts, ps-daemon.ts, wps-client.ts
WPS 客户端与后端通道实现目录：ps-daemon / wps-client / action-routes 走 Windows COM 桥；
addon-channel.ts 是 Linux/macOS 加载项通道的 Node 侧（127.0.0.1:27740 长轮询，B2 骨架，
加载项侧在仓库根 wps-multitool-addon-linux/）。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。
