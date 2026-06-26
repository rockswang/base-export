<script setup>
import { onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { bitable, FieldType } from '@lark-base-open/js-sdk'
import { ElMessage } from 'element-plus'
import { ZipBuilder } from '../utils/zip-builder.js'
import { XlsxBuilder } from '../utils/xlsx-builder.js'

const { t } = useI18n()
const loading = ref(false)
const progress = ref({ show: false, percent: 0, text: '' })
const scope = ref({ loading: true, tableName: '', viewName: '', records: 0, fields: 0, attachmentFields: 0, error: '' })

const ATTACH = FieldType.Attachment
const CONCURRENCY = 10

function fmt(v) {
  if (v === null || v === undefined) return ''
  if (Array.isArray(v)) return v.map(x => x == null ? '' : x.text ?? x.name ?? x).join(', ')
  if (typeof v === 'object') return v.text ?? v.name ?? JSON.stringify(v)
  return String(v)
}

function safename(n) { return n.replace(/[\\/:*?"<>|]/g, '_') }

function uniqueAttachmentName(name, usedNames) {
  const safe = safename(name || 'attachment')
  const dot = safe.lastIndexOf('.')
  const base = dot > 0 ? safe.slice(0, dot) : safe
  const ext = dot > 0 ? safe.slice(dot) : ''
  let out = safe, i = 1
  while (usedNames.has(out.toLowerCase())) out = `${base}-${i++}${ext}`
  usedNames.add(out.toLowerCase())
  return out
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = filename
  document.body.appendChild(a); a.click()
  document.body.removeChild(a)
  setTimeout(() => URL.revokeObjectURL(url), 30000)
}

async function mapLimit(items, limit, fn, onProgress) {
  const results = new Array(items.length)
  let next = 0, done = 0
  const worker = async () => {
    while (next < items.length) {
      const i = next++
      results[i] = await fn(items[i], i)
      done++; onProgress?.(done, items.length)
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, () => worker()))
  return results
}

async function fetchAttachment(url) {
  const resp = await fetch(url)
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`)
  const ctype = (resp.headers.get('content-type') || '').toLowerCase()
  if (ctype.includes('text/html')) {
    const text = await resp.text()
    throw new Error(`Bad response (${ctype}): ${text.substring(0, 200)}`)
  }
  // 注意：application/json 也可能是 JSON 文件本身，直接按二进制读取
  // 飞书下载接口会正确设置 Content-Disposition: attachment
  return resp.arrayBuffer()
}

function createDownloadPool(limit, onProgress) {
  const queue = [], results = []
  let active = 0, total = 0, done = 0, closed = false, resolveIdle
  const idle = new Promise(resolve => { resolveIdle = resolve })
  const settleIfIdle = () => { if (closed && active === 0 && queue.length === 0) resolveIdle(results) }
  const pump = () => {
    while (active < limit && queue.length) {
      const task = queue.shift(); active++
      fetchAttachment(task.url)
        .then(buffer => results.push({ ...task, buffer }))
        .catch(e => { console.warn('[下载失败]', task.name, task.url, e.message); results.push({ ...task, error: e.message }) })
        .finally(() => { active--; done++; onProgress?.(done, total, task.name); pump(); settleIfIdle() })
    }
  }
  return {
    add(task) { total++; queue.push(task); pump(); onProgress?.(done, total, task.name) },
    finish() { closed = true; settleIfIdle(); return idle },
  }
}

async function getScopeInfo() {
  scope.value = { loading: true, tableName: '', viewName: '', records: 0, fields: 0, attachmentFields: 0, error: '' }
  try {
    const sel = await bitable.base.getSelection()
    if (!sel?.tableId || !sel?.viewId) throw new Error(t('error.noView'))
    const table = await bitable.base.getTableById(sel.tableId)
    const view = await table.getViewById(sel.viewId)
    const tableName = await table.getName()
    const viewName = await view.getName()
    const tableFields = await table.getFieldMetaList()
    const visibleFieldIds = await view.getVisibleFieldIdList()
    const byId = new Map(tableFields.map(f => [f.id, f]))
    const fields = visibleFieldIds.map(id => byId.get(id)).filter(Boolean)
    const records = (await view.getVisibleRecordIdList()).filter(Boolean).length
    scope.value = { loading: false, tableName, viewName, records, fields: fields.length, attachmentFields: fields.filter(f => f.type === ATTACH).length, error: '' }
  } catch (e) {
    scope.value = { loading: false, tableName: '', viewName: '', records: 0, fields: 0, attachmentFields: 0, error: e.message }
  }
}

async function doExport() {
  if (loading.value) return
  loading.value = true
  progress.value = { show: true, percent: 0, text: t('progress.readingTable') }
  try {
    const sel = await bitable.base.getSelection()
    if (!sel?.tableId) { ElMessage.error(t('error.noTable')); return }
    if (!sel?.viewId) { ElMessage.error(t('error.noView')); return }
    const table = await bitable.base.getTableById(sel.tableId)
    const view = await table.getViewById(sel.viewId)
    const tableName = await table.getName()

    // ── 字段元数据 ──
    progress.value.text = t('progress.readingFields')
    const tableFields = await table.getFieldMetaList()
    const visibleFieldIds = await view.getVisibleFieldIdList()
    const byId = new Map(tableFields.map(f => [f.id, f]))
    const fields = visibleFieldIds.map(id => byId.get(id)).filter(Boolean)
    const aFields = fields.filter(f => f.type === ATTACH)
    const nFields = fields.filter(f => f.type !== ATTACH)
    progress.value.text = t('progress.fieldSummary', { n: fields.length, m: aFields.length })
    const header = [...nFields.map(f => f.name), ...aFields.map(f => f.name)]

    // ── 获取附件字段实例 ──
    const aFieldInstances = await Promise.all(aFields.map(af => table.getFieldById(af.id)))

    // ── 获取记录列表 ──
    progress.value.text = t('progress.fetchingRecords')
    const recordIds = (await view.getVisibleRecordIdList()).filter(Boolean)
    if (!recordIds.length) { ElMessage.warning(t('error.noData')); return }

    // ── 读取数据 + 通过 SDK 获取附件 URL ──
    progress.value.text = t('progress.processing', { n: recordIds.length })
    const dlIdx = new Map()
    const usedAttachmentNames = new Set()
    const downloadPool = createDownloadPool(CONCURRENCY)
    const rowResults = await mapLimit(recordIds, CONCURRENCY, async (rid) => {
      const vals = {}; for (const f of nFields) vals[f.id] = await table.getCellValue(f.id, rid)

      const nv = nFields.map(f => fmt(vals[f.id]))
      const rows = []
      const sdkUrlFailed = []

      const attCols = []
      for (let ai = 0; ai < aFields.length; ai++) {
        const af = aFields[ai]
        const fi = aFieldInstances[ai]
        const v = await fi.getValue(rid)
        const col = []
        attCols.push(col)
        if (!Array.isArray(v) || !v.length) continue
        try {
          const urls = await fi.getAttachmentUrls(rid)
          for (let i = 0; i < v.length; i++) {
            const path = uniqueAttachmentName(v[i].name, usedAttachmentNames)
            col.push({ name: v[i].name, type: v[i].type, url: urls[i] ?? null, token: v[i].token, path })
          }
        } catch (e) {
          console.warn(`[getAttachmentUrls 失败] ${af.name} rid=${rid}:`, e)
          sdkUrlFailed.push({ field: af.name, recordId: rid, error: e.message })
          for (let i = 0; i < v.length; i++) {
            const path = uniqueAttachmentName(v[i].name, usedAttachmentNames)
            col.push({ name: v[i].name, type: v[i].type, url: null, token: v[i].token, path })
          }
        }
      }

      const rowCount = Math.max(1, ...attCols.map(col => col.length))
      for (let i = 0; i < rowCount; i++) {
        const cells = i === 0 ? [...nv] : nFields.map(() => '')
        for (const col of attCols) {
          const a = col[i]
          if (!a) { cells.push(''); continue }
          if (a.url) {
            cells.push({ _type: 'url', _name: a.name, _url: a.url, _path: `attachments/${a.path}` })
            const key = `${a.url}\n${a.path}`
            if (!dlIdx.has(key)) { dlIdx.set(key, true); downloadPool.add({ name: a.name, type: a.type, url: a.url, path: a.path }) }
          } else {
            cells.push(t('cell.attachment', { name: a.name, token: a.token }))
          }
        }
        rows.push({ cells, height: 22 })
      }

      return { rows, sdkUrlFailed }
    }, (done, total) => {
      progress.value.percent = Math.min(85, Math.round(done / total * 85))
      progress.value.text = t('progress.processed', { done, total })
    })

    const rows = []
    const sdkUrlFailed = []
    for (const result of rowResults) {
      rows.push(...result.rows)
      sdkUrlFailed.push(...result.sdkUrlFailed)
    }

    if (sdkUrlFailed.length) {
      ElMessage.warning(t('warning.urlFailed', { n: sdkUrlFailed.length }))
      console.warn('SDK URL 失败明细:', sdkUrlFailed.slice(0, 5))
    }

    const results = await downloadPool.finish()
    if (results.length) {
      progress.value.text = t('progress.downloadDone')

      const failed = []
      for (const r of results) { if (r.error) failed.push(r.name) }
      if (failed.length) ElMessage.warning(t('warning.downloadFailed', { n: failed.length }))

      progress.value.text = t('progress.buildingXlsx'); progress.value.percent = 90
      const xb = await new XlsxBuilder().init()
      const xlsxBuf = xb.build({ rows: { header, data: rows } })

      progress.value.text = t('progress.buildingZip'); progress.value.percent = 95
      const zb = new ZipBuilder()
      zb.addXlsx('导出数据.xlsx', new Uint8Array(xlsxBuf))
      for (const r of results) { if (r.buffer) zb.addAttachment(r.path, new Uint8Array(r.buffer)) }

      triggerDownload(await zb.generate(), `${safename(tableName || 'export')}.zip`)
      progress.value.text = t('progress.done', { n: recordIds.length, m: results.length })
    } else {
      progress.value.text = t('progress.buildingXlsx'); progress.value.percent = 90
      const xb = await new XlsxBuilder().init()
      const xlsxBuf = xb.build({ rows: { header, data: rows } })
      const zb = new ZipBuilder()
      zb.addXlsx('导出数据.xlsx', new Uint8Array(xlsxBuf))
      triggerDownload(await zb.generate(), `${safename(tableName || 'export')}.zip`)
      progress.value.text = t('progress.doneNoAttachment', { n: recordIds.length })
    }
    progress.value.percent = 100
  } catch (e) {
    console.error(e)
    ElMessage.error(t('error.exportFailed', { msg: e.message }))
    progress.value.text = t('progress.failed', { msg: e.message })
  } finally { loading.value = false }
}

onMounted(getScopeInfo)
</script>

<template>
  <div class="panel">
    <h3><img src="/icon.png" alt="" class="plugin-icon">{{ $t('panel.title') }}</h3>
    <el-button type="primary" size="large" :loading="loading" style="width:100%" @click="doExport">
      {{ loading ? $t('button.exporting') : $t('button.export') }}
    </el-button>
    <div class="help">
      <p>{{ $t('help.line1') }}</p>
      <p v-html="$t('help.line2', { xlsx: '<span>' + $t('help.xlsx') + '</span>', attachments: '<span>' + $t('help.attachments') + '</span>' })"></p>
    </div>
    <div class="scope">
      <template v-if="scope.loading">{{ $t('scope.loading') }}</template>
      <template v-else-if="scope.error">{{ scope.error }}</template>
      <template v-else>
        <div>{{ $t('scope.table', { name: scope.tableName }) }}</div>
        <div>{{ $t('scope.view', { name: scope.viewName }) }}</div>
        <div>{{ $t('scope.records', { n: scope.records }) }}</div>
        <div>{{ $t('scope.fields', { n: scope.fields, m: scope.attachmentFields }) }}</div>
      </template>
    </div>
    <div v-if="progress.show" class="progress">
      <el-progress :percentage="progress.percent" :status="progress.percent >= 100 ? 'success' : undefined" />
      <p class="tip">{{ progress.text }}</p>
    </div>
  </div>
</template>

<style scoped>
.panel { padding: 16px; }
.panel h3 { margin-bottom: 16px; font-size: 22px; display: flex; align-items: center; gap: 10px; }
.plugin-icon { width: 56px; height: 56px; border-radius: 6px; }
.help { margin-top: 12px; font-size: 13px; line-height: 1.6; color: #646a73; }
.help p { margin: 0 0 4px; }
.help span { color: #1f2329; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
.scope { margin-top: 12px; padding: 10px 12px; border: 1px solid #dee0e3; border-radius: 6px; font-size: 13px; line-height: 1.7; color: #1f2329; background: #f7f8fa; }
.progress { margin-top: 16px; }
.tip { margin-top: 8px; font-size: 13px; color: #666; word-break: break-all; }
</style>