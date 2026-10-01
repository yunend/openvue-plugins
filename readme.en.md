# openvue-plugins

[English](readme.en.md) | [**中文**](README.md)

The plugin repository for the openvue desktop app. It currently bundles two Office preview plugins: `office-vue` and `office-easy`.

## Structure

Each plugin is a standalone directory containing two required parts:

- `plugin.json`: plugin manifest (id, version, supported extensions, etc.)
- `dist-plugin/`: build output (pointed to by Vite's `build.outDir`)

## Plugin Development

Take a Vue plugin as an example (refer to `office-vue`):

```bash
mkdir my-plugin && cd my-plugin
npm create vite@latest . -- --template vue
npm install
```

`vite.config.js` outputs the build to `dist-plugin/` (`base: './'`):

```js
export default defineConfig({
  plugins: [vue()],
  base: './',
  build: { outDir: './dist-plugin' },
})
```

Create `plugin.json` (the id must be globally unique; release fields such as `sha256`/`sizeBytes` can be left empty):

```json
{
  "id": "my-plugin",
  "name": "My Plugin",
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

Local development and build:

```bash
npm run dev      # local preview
npm run build    # build to dist-plugin/
```

## Updates and Publishing

Plugin update logic runs automatically in GitHub Actions (`.github/workflows/release.yml`), triggered by a `v*` tag:

1. Package each plugin's `dist-plugin/` into `{id}-{version}.zip` and compute its `sha256`.
2. Compare the hash against the previous `plugins-index.json` on the `plugins-data` branch:
   - **Hash unchanged** → skip publishing and reuse the old download URL.
   - **Hash changed** → publish a new ZIP and write back `sha256`, `sizeBytes`, `publishedAt`, and `downloadSources`.
3. Aggregate and generate a new `plugins-index.json`, create a GitHub Release, and force-push it to the `plugins-data` branch for the main app to query updates.

Update steps:

```bash
npm run build                      # 1. rebuild
# 2. commit (including dist-plugin and plugin.json) and push
# 3. tag to trigger automated CI release
# 4. skip automatically if the output is unchanged
```

## Notes

- `dist-plugin/` must be committed so that CI can read and package it.
- Packaging uses `zip -X` with a unified timestamp to guarantee identical hashes for identical content, which is what enables "skip re-publishing when nothing changed".
- Plugin ids should not be changed after release, otherwise the plugin will be treated as a new one.
- The `plugins-data` branch only stores the latest index (force-overwritten, including `urlTemplate` addresses).

## License

[MIT](LICENSE)