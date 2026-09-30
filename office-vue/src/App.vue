<script setup>
import { ref, computed, onMounted, watch, nextTick, defineAsyncComponent } from 'vue'

const DocxPreview = defineAsyncComponent(() =>
  Promise.all([
    import('@vue-office/docx/lib/index.css'),
    import('@vue-office/docx')
  ]).then(([, mod]) => mod)
)

const ExcelPreview = defineAsyncComponent(() =>
  Promise.all([
    import('@vue-office/excel/lib/index.css'),
    import('@vue-office/excel')
  ]).then(([, mod]) => mod)
)

const PdfPreview = defineAsyncComponent(() =>
  Promise.all([
    import('@vue-office/pdf')
  ]).then(([mod]) => mod)
)

// 支持的文件后缀 -> 类型映射
const SUPPORTED_EXT = {
  docx: 'word',
  doc: 'word',
  xlsx: 'excel',
  xls: 'excel',
  pdf: 'pdf'
}

const filePath = ref('')
const loading = ref(false)
const errorMsg = ref('')
const fileType = ref('')
// 预览视图缩放比例（1 = 100%）
const zoom = ref(1)
const ZOOM_MIN = 0.5
const ZOOM_MAX = 3
const ZOOM_STEP = 0.2

// 放大
const zoomIn = () => {
  zoom.value = Math.min(ZOOM_MAX, Math.round((zoom.value + ZOOM_STEP) * 100) / 100)
}

// 缩小
const zoomOut = () => {
  zoom.value = Math.max(ZOOM_MIN, Math.round((zoom.value - ZOOM_STEP) * 100) / 100)
}

// 重置缩放
const zoomReset = () => {
  zoom.value = 1
}

// PDF 静态资源（cmaps）路径，指向本地，避免从 unpkg CDN 加载
const pdfStaticFileUrl = import.meta.env.BASE_URL

// Excel 预览配置
const excelOptions = {
  xls: false,
  minColLength: 0,
  minRowLength: 0
}

// PDF 预览配置
const pdfOptions = {
  gap: 10
}

// 从 URL 获取 path 参数（手动解析，避免 URLSearchParams 把 + 解码为空格）
const getQueryParam = (name) => {
  const search = window.location.search
  if (!search || search.length <= 1) return ''
  const params = search.substring(1).split('&')
  for (const param of params) {
    const eqIndex = param.indexOf('=')
    if (eqIndex === -1) continue
    const key = decodeURIComponent(param.substring(0, eqIndex))
    if (key === name) {
      return decodeURIComponent(param.substring(eqIndex + 1))
    }
  }
  return ''
}

// 获取文件后缀（小写，不含点）
const getExtension = (path) => {
  if (!path) return ''
  // 去除查询参数后再取后缀
  const cleanPath = path.split('?')[0].split('#')[0]
  const match = cleanPath.match(/\.([^.\/\\]+)$/)
  return match ? match[1].toLowerCase() : ''
}

// 文件名显示
const fileName = computed(() => {
  if (!filePath.value) return ''
  const cleanPath = filePath.value.split('?')[0].split('#')[0]
  const parts = cleanPath.split(/[\/\\]/)
  return parts[parts.length - 1] || cleanPath
})

// 支持的后缀列表显示
const supportedExtList = Object.keys(SUPPORTED_EXT).map(e => `.${e}`).join(', ')

// 页面 URL origin（在 onMounted 中赋值，避免模板直接访问 window）
const pageOrigin = ref('http://localhost:5173')

// 加载并预览文件
const loadFile = async (path) => {
  if (!path) {
    fileType.value = ''
    errorMsg.value = ''
    return
  }

  loading.value = true
  errorMsg.value = ''

  await nextTick()

  try {
    const ext = getExtension(path)
    const type = SUPPORTED_EXT[ext]

    if (!type) {
      fileType.value = ''
      errorMsg.value = `不支持的文件格式: .${ext || '(未知)'}`
      loading.value = false
      return
    }

    fileType.value = type
  } catch (err) {
    loading.value = false
    errorMsg.value = `加载文件失败: ${err.message || err}`
  }
}

onMounted(() => {
  // 初始化 origin
  if (typeof window !== 'undefined' && window.location) {
    pageOrigin.value = window.location.origin
  }
  const path = getQueryParam('path')
  filePath.value = path
  if (path) {
    loadFile(path)
  }
  // 监听 URL 变化
  window.addEventListener('popstate', onUrlChange)
})

// 监听 URL hash/query 变化（简单的 popstate 监听）
const onUrlChange = () => {
  const path = getQueryParam('path')
  if (path !== filePath.value) {
    filePath.value = path
    loadFile(path)
  }
}

// vue-office 事件处理
const onDocxRendered = () => { loading.value = false }
const onExcelRendered = () => { loading.value = false }
const onPdfRendered = () => { loading.value = false }
const onPreviewError = (err) => {
  loading.value = false
  errorMsg.value = `预览出错: ${err?.message || err || '未知错误'}`
}
</script>

<template>
  <div class="preview-container">
    <!-- 顶部工具栏 -->
    <header class="toolbar">
      <div class="toolbar-title">
        <span class="icon">📄</span>
        <span>Office 文件预览</span>
      </div>
      <div class="toolbar-actions">
        <div v-if="fileType" class="zoom-control">
          <button class="action-btn" title="缩小" @click="zoomOut">－</button>
          <button class="zoom-value" title="重置缩放" @click="zoomReset">{{ Math.round(zoom * 100) }}%</button>
          <button class="action-btn" title="放大" @click="zoomIn">＋</button>
        </div>
      </div>
    </header>


    <!-- 加载中（尚未确定文件类型） -->
    <div v-if="loading && !fileType" class="status-box loading">
      <div class="spinner"></div>
      <span>加载中...</span>
    </div>

    <!-- 错误信息 -->
    <div v-else-if="errorMsg" class="status-box error">
      <span class="status-icon">⚠️</span>
      <span>{{ errorMsg }}</span>
      <br />
      <small>支持的格式: {{ supportedExtList }}</small>
    </div>

    <!-- 无文件 -->
    <div v-else-if="!filePath" class="status-box empty">
      <span class="status-icon">📂</span>
      <h2>请指定要预览的文件</h2>
      <p>在 URL 中添加 <code>?path=文件的URL路径</code> 参数</p>
      <p>支持的格式: <code>{{ supportedExtList }}</code></p>
      <div class="example">
        <p>示例：</p>
        <code>{{ pageOrigin }}?path=https://example.com/demo.docx</code>
      </div>
    </div>

    <!-- 预览区域（始终渲染，loading 作为遮罩覆盖在上面） -->
    <div v-else-if="fileType" class="preview-wrapper" :class="`preview-${fileType}`" :style="{ '--zoom': zoom }">
      <!-- 加载中遮罩 -->
      <div v-if="loading" class="loading-overlay">
        <div class="spinner"></div>
        <span>加载中...</span>
      </div>
      <div class="zoom-content" :style="{ zoom }">
      <!-- Word 预览 -->
      <DocxPreview
        v-if="fileType === 'word'"
        :key="filePath"
        :src="filePath"
        style="width: 100%;"
        @rendered="onDocxRendered"
        @error="onPreviewError"
      />
      <!-- Excel 预览 -->
      <ExcelPreview
        v-if="fileType === 'excel'"
        :key="filePath"
        :src="filePath"
        :options="excelOptions"
        style="height: 100%;"
        @rendered="onExcelRendered"
        @error="onPreviewError"
      />
      <!-- PDF 预览 -->
      <PdfPreview
        v-if="fileType === 'pdf'"
        :key="filePath"
        :src="filePath"
        :staticFileUrl="pdfStaticFileUrl"
        :options="pdfOptions"
        @rendered="onPdfRendered"
        @error="onPreviewError"
      />
      </div>
    </div>
  </div>
</template>

<style scoped>
.preview-container {
  height: 100vh;            /* 确定高度：打通 flex 高度链 */
  display: flex;
  flex-direction: column;
  background: #f5f7fa;
}

.toolbar {
  background: #fff;
  padding: 12px 24px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
  display: flex;
  align-items: center;
  gap: 24px;
  flex-wrap: wrap;
  position: sticky;
  top: 0;
  z-index: 10;
}

.toolbar-title {
  font-size: 18px;
  font-weight: 600;
  color: #1f2937;
  display: flex;
  align-items: center;
  gap: 8px;
  white-space: nowrap;
}

.toolbar-title .icon { font-size: 22px; }

.toolbar-actions {
  margin-left: auto;
  display: flex;
  align-items: center;
}

.zoom-control {
  display: flex;
  align-items: center;
  gap: 4px;
  background: #f3f4f6;
  border-radius: 8px;
  padding: 4px;
}

.zoom-control button {
  border: none;
  background: transparent;
  cursor: pointer;
  font: inherit;
  color: #4b5563;
  display: flex;
  align-items: center;
  justify-content: center;
  height: 28px;
  border-radius: 6px;
  transition: background 0.2s;
}

.zoom-control .action-btn { width: 28px; font-size: 16px; }
.zoom-control .zoom-value {
  min-width: 52px;
  padding: 0 8px;
  font-size: 13px;
  font-weight: 600;
  color: #374151;
}
.zoom-control .action-btn:hover,
.zoom-control .zoom-value:hover { background: #e5e7eb; }
.zoom-control .action-btn:hover { color: #1f2937; }

.status-box {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 60px 24px;
  text-align: center;
  color: #6b7280;
}

.status-box .status-icon {
  font-size: 56px;
  margin-bottom: 20px;
}

.status-box h2 {
  margin: 0 0 12px;
  color: #1f2937;
  font-size: 22px;
}

.status-box p {
  margin: 8px 0;
  font-size: 14px;
}

.status-box code {
  background: #f3f4f6;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 13px;
  color: #1f2937;
}

.status-box.loading {
  flex-direction: row;
  gap: 14px;
  font-size: 15px;
  color: #3b82f6;
}

.status-box.error { color: #b91c1c; }
.status-box.error small { color: #6b7280; margin-top: 12px; }

.example {
  margin-top: 24px;
  padding: 16px 20px;
  background: #f9fafb;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  max-width: 600px;
  text-align: left;
}

.example p {
  margin: 0 0 8px;
  font-weight: 500;
  color: #374151;
}

.example code {
  display: block;
  background: #fff;
  padding: 10px 12px;
  border: 1px solid #e5e7eb;
  word-break: break-all;
}

.spinner {
  width: 24px;
  height: 24px;
  border: 3px solid #dbeafe;
  border-top-color: #3b82f6;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin { to { transform: rotate(360deg); } }

.preview-wrapper {
  flex: 1;
  min-height: 0;   /* flex 子元素需要这个才能正确 overflow */
  padding: 16px 24px 24px;
  overflow: auto;
  position: relative;
}

.loading-overlay {
  position: absolute;
  inset: 0;
  background: rgba(245, 247, 250, 0.92);
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 16px;
  z-index: 20;
  font-size: 15px;
  color: #3b82f6;
}

/* 缩放内容区：CSS zoom 会真实影响布局尺寸，父容器可正确撑开 */
.zoom-content { width: 100%; min-height: 0; }

/* vue-office 组件通用宽度 */
.preview-wrapper :deep(.vue-office-docx),
.preview-wrapper :deep(.vue-office-excel),
.preview-wrapper :deep(.vue-office-pdf) { width: 100%; }

/* === Excel（x-spreadsheet 虚拟滚动需要整条容器链确定 height） === */
.preview-wrapper.preview-excel .zoom-content {
  height: 100%;
  min-height: calc(100% * (1 / var(--zoom, 1)));   /* 补偿 zoom 缩放 */
}
.preview-wrapper.preview-excel :deep(.vue-office-excel),
.preview-wrapper.preview-excel :deep(.x-spreadsheet),
.preview-wrapper.preview-excel :deep(.x-spreadsheet-sheet) {
  height: 100% !important;
  min-height: 100%;
}

/* === PDF（组件自身处理滚动） === */
.preview-wrapper.preview-pdf { overflow: hidden; }
.preview-wrapper.preview-pdf .zoom-content { height: 100%; overflow: hidden; }
.preview-wrapper.preview-pdf :deep(.vue-office-pdf) {
  height: 100%;
  overflow-y: auto !important;
}

/* 通用表格（docx 等原生 table） */
.preview-wrapper :deep(table) { width: 100%; border-collapse: collapse; }
.preview-wrapper :deep(td),
.preview-wrapper :deep(th) {
  border: 1px solid #d1d5db;
  padding: 6px 10px;
  white-space: nowrap;
}
.preview-wrapper :deep(th) { background: #f3f4f6; font-weight: 600; }
</style>