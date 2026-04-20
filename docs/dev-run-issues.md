# Word-Cursor 运行问题记录

日期：2026-02-24
命令：`npm run dev:electron`

---

## 1. `concurrently` 命令未找到

**现象**：运行 `npm run dev:electron` 时报错 `'concurrently' 不是内部或外部命令`。

**原因**：`node_modules` 未安装，`concurrently` 是 devDependency，需要先执行 `npm install`。

**解决**：执行 `npm install` 安装全部依赖。

---

## 2. npm install 依赖废弃警告

**现象**：安装过程中出现多个 deprecated 警告：
- `inflight@1.0.6` — 内存泄漏，建议用 `lru-cache` 替代
- `npmlog@6.0.2`、`are-we-there-yet@3.0.1`、`gauge@4.0.4` — 不再维护
- `rimraf@2.7.1` / `rimraf@3.0.2` — v4 以下不再支持
- `glob@7.2.3` / `glob@8.1.0` — v9 以下不再支持
- `fstream@1.0.12`、`boolean@3.2.0`、`@npmcli/move-file@2.0.1` — 不再支持
- `lodash.isequal@4.5.0` — 建议用 `node:util.isDeepStrictEqual` 替代

**影响**：不影响运行，但存在 43 个安全漏洞（6 moderate, 37 high）。

**建议**：运行 `npm audit fix` 处理可自动修复的漏洞，剩余需评估是否升级相关依赖。

---

## 3. 端口 3000 被占用

**现象**：Vite 启动失败，报错 `Error: Port 3000 is already in use`。因为 `--strictPort` 参数，端口被占用时直接退出而非自动切换。

**解决**：找到占用端口 3000 的进程（PID 50724）并终止后重新启动。

**排查命令**：
```bash
netstat -ano | grep ':3000'
taskkill //PID <PID> //F
```

---

## 4. 文件服务器端口连续被占用

**现象**：Electron 内置文件服务器默认端口 9090 被占用，自动尝试 9091 也被占用，最终使用 9092 启动。

```
文件服务器端口 9090 被占用，尝试端口 9091...
文件服务器端口 9091 被占用，尝试端口 9092...
📁 本地文件服务器已启动: http://localhost:9092
```

**影响**：有自动递增逻辑，不影响功能，但说明本机有其他服务占用了 9090-9091。

---

## 5. Electron GPU 缓存创建失败

**现象**：启动时连续报错：
```
ERROR:cache_util_win.cc(20) Unable to move the cache: 拒绝访问。(0x5)
ERROR:disk_cache.cc(208) Unable to create cache
ERROR:gpu_disk_cache.cc(713) Gpu Cache Creation failed: -2
```

**原因**：Electron/Chromium 的 GPU 磁盘缓存目录权限不足或被其他进程锁定（常见于 Windows 上多个 Electron 实例同时运行）。

**影响**：不影响应用核心功能，GPU 加速可能降级为软件渲染。

**建议**：关闭其他 Electron 应用后重启，或清理 `%APPDATA%/word-cursor/GPUCache` 目录。

---

## 6. Service Worker 数据库 IO 错误

**现象**：
```
ERROR:service_worker_storage.cc(2072) Failed to delete the database: Database IO error
```

**原因**：与上述缓存权限问题相关，Service Worker 存储无法正常清理。

**影响**：不影响主要功能，可能导致旧缓存残留。

---

## 7. DevTools 控制台警告

**现象**：
- `Unknown VE context: language-mismatch` — Chromium DevTools 内部的 visual logging 警告
- `Request Autofill.enable failed` / `Request Autofill.setAddresses failed` — Autofill CDP 协议方法不存在

**影响**：纯 DevTools 内部问题，不影响应用功能，可忽略。

---

## 8. baseline-browser-mapping 数据过期

**现象**：
```
[baseline-browser-mapping] The data in this module is over two months old.
```

**建议**：执行 `npm i baseline-browser-mapping@latest -D` 更新。

---

## 总结

| 问题 | 严重程度 | 是否阻塞启动 |
|------|---------|-------------|
| concurrently 未安装 | 高 | 是（需先 npm install） |
| 依赖废弃 + 安全漏洞 | 中 | 否 |
| 端口 3000 被占用 | 高 | 是（需手动释放端口） |
| 文件服务器端口递增 | 低 | 否（自动处理） |
| GPU 缓存权限拒绝 | 低 | 否 |
| Service Worker IO 错误 | 低 | 否 |
| DevTools 警告 | 无 | 否 |
| baseline-browser-mapping 过期 | 低 | 否 |
