# 浏览器端可编辑 PPTX 导出

## 状态与目标

第一阶段已实现并完成 Web 部署试点。实现基于上游 revision
`73953213a6fec2c8092e8e77d229a3074aa828a9`，复用该版本桌面端的
`dom-to-pptx` 2.0.1 browser UMD 和 DOM 规范化算法。

当 `/api/version` 返回 `capabilities.slideRenderer=false` 时，Chrome/Edge
可通过「导出 → PPTX → 浏览器可编辑」下载文件，转换完全在用户浏览器中执行。
服务端继续提供现有 HTML 和静态资产，不新增渲染 API、Electron、Chromium 或 CDP。
桌面端及具备服务端渲染器的环境保留原有可编辑和截图两种导出模式。

## 可编辑边界

| 内容 | 导出形态 |
| --- | --- |
| 标题、正文、列表 | PowerPoint 原生文本框 |
| 基础形状、纯色填充 | 原生 shape |
| SVG | 矢量图对象，支持整体移动缩放，不承诺路径节点编辑 |
| 图片 | 嵌入位图 |
| 复杂渐变、滤镜、混合背景 | 栅格化或近似降级，并提示警告 |
| 内嵌 iframe | 不支持，省略并提示警告 |

不承诺复杂视觉效果与 Electron 导出像素级一致。批量、无人值守及高保真导出需要
另行设计服务端渲染方案，不属于当前实现。

## 实现结构

- `apps/web/scripts/build-client-pptx.ts`：解压已纳入版本控制的桌面端 UMD，
  校验 SHA-256，并编译 iframe bridge。由 Web `dev` 和 `build` 自动调用。
- `client-pptx-protocol.ts`：消息类型、关联校验、大小限制、能力路由及错误 URL 脱敏。
- `clientPptxExport.ts`：同源加载并校验脚本、创建独立 iframe、握手、下载和清理。
- `pptx-export-bridge.ts`：在 iframe 内等待资源和布局、处理图片并调用转换器。
- `pptx-export-normalizer.ts`：从指定 revision 移植 DOM 规范化算法，处理 authored size、
  slide 可见性、背景、标题换行、CJK 字体和 SVG className。
- `srcdoc.ts`：复用现有 HTML 注入点和跳转保护，将 UMD 与 bridge 注入作者 CSP 之前。
- `FileViewer.tsx`：复用 source HTML、相对资产处理、scoped preview base、埋点和通知；
  无服务端渲染器时仅提供浏览器可编辑模式。选项文案覆盖全部 19 种语言。
- `exports.ts`：并列导出客户端入口，保留 `exportProjectAsPptx(opts)` 的签名及实现。

除构建脚本外，上述运行时文件均位于 `apps/web/src/runtime/`。

## 隔离与生命周期

导出使用独立 iframe，sandbox 固定为 `allow-scripts allow-downloads`，不加入
`allow-same-origin`，也不读取或改写 live preview 的 DOM。

父页面同源读取带版本标记的 UMD 和 bridge 文本后内联进 srcdoc，避免 opaque-origin
iframe 请求脚本时的认证限制。bridge 安装监听器后发送 ready，父页面校验消息来源并
传入专用 MessageChannel。实际转换等待 load、字体、图片解码和布局完成。

离屏 iframe 的动画帧可能被 Chromium 节流，因此两帧布局等待另有 100ms 计时兜底。
请求标识通过 `crypto.getRandomValues` 生成 128-bit 随机值，兼容 HTTP 部署。
结果使用 transferable ArrayBuffer 传递，并检查请求标识、schema、大小和 ZIP 文件头。

MessageChannel 防止其他窗口串扰，但不隔离同一 iframe JS world 中的作者脚本。
安全边界仍是 sandbox、最小权限的 scoped asset base 及输入输出限制。

全程单任务互斥；超时、取消、异常和成功均清理 iframe、端口、监听器与 Object URL。
关键图片或声明字体加载失败时整体失败，错误 URL 脱敏，不静默生成缺失资源的文件。

## 构建与开关

```sh
pnpm install --frozen-lockfile
pnpm --filter @open-design/web build
```

不要额外安装 `dom-to-pptx` npm 包，其依赖树包含 Puppeteer。固定使用
`apps/desktop/vendor/dom-to-pptx/dom-to-pptx.bundle.js.gz`；解压后 SHA-256 为
`0308535fd30c30fe78df78ed7d45f2d24e7393285bf09acbb2090189b4d50558`。

以下文件在构建时生成并被 Git 忽略：

- `apps/web/public/vendor/dom-to-pptx.bundle.js`
- `apps/web/public/vendor/pptx-export-bridge.js`
- `apps/web/public/vendor/client-pptx.json`

将部署产物中 `vendor/client-pptx.json` 的 `enabled` 改为 `false` 可以立即阻止新的
浏览器导出；刷新后入口隐藏。保留 `version` 字段。父端每次导出前以 `no-store`
重新读取开关。重新构建默认生成 `enabled:true`，发布时应核对开关状态。

部署需整体更新 Web 静态产物，单独复制 vendor 文件不会更新 UI 和协调器。
保留原静态产物或原镜像即可回滚；静态覆盖挂载方案需在每次容器重建时保留其附加配置。
运行数据目录遵循根 `AGENTS.md` 的 Daemon data directory contract。

## 限制与验证

当前保护阈值为 120s 全程超时、100 页、20 MiB HTML、50,000 个 slide 子元素、
单个远程图片 20 MiB、远程图片累计 96 MiB，以及 128 MiB PPTX 结果。
超过 40 页或输出超过 32 MiB 会提示警告。这些阈值不代表已完成相应规模的压测。

已完成的试点验证：

- 全仓 `pnpm guard`、`pnpm typecheck` 及 Web 正式构建通过。
- 导出 runtime 测试 106 项通过，相关 FileViewer 组件测试 9 项通过。
- Edge 中现有 10 页中文项目成功下载；ZIP 包含 62 个原生文本节点，尺寸为 16:9。
  正式 UI 导出约 2.3s，输出约 49 KB，没有请求服务端 `/export/pptx`。
- 补充样本覆盖四种 slide selector、缩略图过滤、中英文、PNG、SVG 和 4:3 尺寸。
- 关键图片缺失会报错且清理 iframe；单元测试覆盖来源伪造、请求关联、超时、取消、
  互斥、运行时禁用和 HTTP 环境兼容。
- 实际导出前后预览元素、`src/srcdoc`、滚动位置和当前页保持不变；主预览截图逐像素一致。

尚未完成：PowerPoint/WPS 中实际修改文本后保存重开、远程字体嵌入的代表性样本、
复杂效果保真度和大规模导出的峰值内存测试。系统中文字体样本通过，不等同于所有
远程 CJK 子集字体均已验证。Firefox/Safari 不在首期支持范围内。
