# openvue-plugins

openvue 桌面应用的插件库，目前内置 `office-vue`、`office-easy` 两个 Office 预览插件。

## 结构

每个插件是一个独立目录，含两个必要部分：

- `plugin.json`：插件清单（id、版本、支持扩展名等）
- `dist-plugin/`：构建产物（Vite `build.outDir` 指向这里）

## 插件开发

以 Vue 插件为例（参考 `office-vue`）：

```bash
mkdir my-plugin && cd my-plugin
npm create vite@latest . -- --template vue
npm install
```

`vite.config.js` 将产物输出到 `dist-plugin/`（`base: './'`）：

```js
export default defineConfig({
  plugins: [vue()],
  base: './',
  build: { outDir: './dist-plugin' },
})
```

创建 `plugin.json`（id 全局唯一，`sha256`/`sizeBytes` 等发布字段留空即可）：

```json
{
  "id": "my-plugin",
  "name": "我的插件",
  "version": "0.1.0",
  "extensions": ["txt"],
  "urlTemplate": "/plugins/{pluginId}/?path={publicPath}",
  "downloadSources": [],
  "sha256": "",
  "archiveFormat": "zip",
  "sizeBytes": 0,
  "publishedAt": ""
}
```

本地开发与构建：

```bash
npm run dev      # 本地预览
npm run build    # 构建到 dist-plugin/
```

## 更新与发布

插件更新逻辑在 GitHub Actions（`.github/workflows/release.yml`）中自动完成，打 `v*` 标签即触发：

1. 将每个插件的 `dist-plugin/` 打成 `{id}-{version}.zip`，计算 `sha256`。
2. 与 `plugins-data` 分支上上次的 `plugins-index.json` 比对哈希：
   - **哈希未变** → 跳过发布，复用旧下载地址。
   - **哈希变化** → 发布新 ZIP，回写 `sha256`、`sizeBytes`、`publishedAt`、`downloadSources`。
3. 汇总生成新的 `plugins-index.json`，创建 GitHub Release，并强制推送到 `plugins-data` 分支供主程序查询更新。

更新步骤：

```bash
npm run build                      # 1. 重新构建
# 2. 提交（含 dist-plugin 与 plugin.json）并推送
# 3. 打 tag 触发 CI 自动发布
# 4. 产物未变化则自动跳过
```

## 注意事项

- `dist-plugin/` 需提交，供 CI 读取打包。
- 打包用 `zip -X` 并统一时间戳，保证相同内容哈希一致，才能实现“无变化不重发”。
- 插件 id 一经发布不宜修改，否则会被当作新插件。
- `plugins-data` 分支仅存最新索引（强制覆盖，含 `urlTemplate` 地址）。

## License

[MIT](LICENSE)