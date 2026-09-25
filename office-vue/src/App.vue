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
    <div v-else-if="fileType" class="preview-wrapper" :class="`preview-${fileType}`">
      <!-- 加载中遮罩 -->
      <div v-if="loading" class="loading-overlay">
        <div class="spinner"></div>
        <span>加载中...</span>
      </div>
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
</template>

<style scoped>
.preview-container {
  min-height: 100vh;
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

.toolbar-title .icon {
  font-size: 22px;
}

.file-type {
  font-size: 11px;
  font-weight: 700;
  padding: 3px 10px;
  border-radius: 4px;
  color: #fff;
  letter-spacing: 0.5px;
}

.file-type.word { background: #2b579a; }
.file-type.excel { background: #217346; }
.file-type.pdf { background: #b91c1c; }

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

.status-box.error {
  color: #b91c1c;
}

.status-box.error small {
  color: #6b7280;
  margin-top: 12px;
}

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

@keyframes spin {
  to { transform: rotate(360deg); }
}

.preview-wrapper {
  flex: 1;
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

/* 让 vue-office 组件填满容器 */
.preview-wrapper > :deep(.vue-office-docx),
.preview-wrapper > :deep(.vue-office-excel),
.preview-wrapper > :deep(.vue-office-pdf) {
  width: 100%;
}

/* PDF 预览修复：让 PDF 组件自身处理滚动，确保虚拟滚动 onScroll 事件能正确触发 */
.preview-wrapper.preview-pdf {
  overflow: hidden;
}

.preview-wrapper.preview-pdf :deep(.vue-office-pdf) {
  height: 100%;
  overflow-y: auto !important;
}

/* Excel 预览表格样式修复：确保表格内容完全展开 */
.preview-wrapper :deep(.vue-office-excel) {
  height: auto !important;
  min-height: 100%;
}

.preview-wrapper :deep(table) {
  width: 100%;
  border-collapse: collapse;
}

.preview-wrapper :deep(td),
.preview-wrapper :deep(th) {
  border: 1px solid #d1d5db;
  padding: 6px 10px;
  white-space: nowrap;
}

.preview-wrapper :deep(th) {
  background: #f3f4f6;
  font-weight: 600;
}
</style>