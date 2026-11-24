/**
 * @module xlsx_lite
 * @path ~/earn/audit-kit/xlsx-lite.mjs
 * @description 零依赖读取 .xlsx（本质是 zip + XML）。只用 Node 内置 zlib，不引入任何第三方包——
 *   保持「装了 Node 就能跑」这个对陌生人的最大优势。
 *   局限：只读第一个工作表；不处理加密、共享公式、图表；富文本只取纯文本拼接。
 * @zone parse
 * @tier Tier1-delivery
 *
 * @skill readXlsx — Buffer → {headers, rows}，与 parseCsv 返回同形
 * @skill columnIndex — 列字母（A/AA）→ 0 基下标
 *
 * @rule fail_loudly_on_unknown_structure P0 — 找不到 zip 中央目录或无工作表 ⇒ 抛错，不返回空表
 *   @enforced-by audit.mjs:--self-test
 * @rule no_silent_cell_loss P1 — 跳过的空单元格用空串补齐，保证列对齐；行内错位即抛错
 *   @enforced-by audit.mjs:--self-test
 */

import { inflateRawSync } from 'node:zlib'

/** XLSX 结构不符预期时抛出。 */
export class XlsxError extends Error {}

/** 中央目录条目签名。 */
const CD_SIG = 0x02014b50
/** 本地文件头签名。 */
const LFH_SIG = 0x04034b50
/** 中央目录结束签名。 */
const EOCD_SIG = 0x06054b50

/**
 * 从 zip 字节里抽出指定文件的内容。
 * @param {Buffer} buffer 整个 zip
 * @returns {Map<string, Buffer>} 文件名 → 解压后内容
 */
function unzip(buffer) {
  let eocd = -1
  const stop = Math.max(0, buffer.length - 65557)
  for (let index = buffer.length - 22; index >= stop; index -= 1) {
    if (buffer.readUInt32LE(index) === EOCD_SIG) { eocd = index; break }
  }
  if (eocd < 0) throw new XlsxError('不是有效的 xlsx：找不到 zip 中央目录（EOCD）')
  const count = buffer.readUInt16LE(eocd + 10)
  let offset = buffer.readUInt32LE(eocd + 16)
  const files = new Map()
  for (let index = 0; index < count; index += 1) {
    if (buffer.readUInt32LE(offset) !== CD_SIG) throw new XlsxError('zip 中央目录条目损坏')
    const method = buffer.readUInt16LE(offset + 10)
    const compressedSize = buffer.readUInt32LE(offset + 20)
    const nameLength = buffer.readUInt16LE(offset + 28)
    const extraLength = buffer.readUInt16LE(offset + 30)
    const commentLength = buffer.readUInt16LE(offset + 32)
    const localOffset = buffer.readUInt32LE(offset + 42)
    const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength)
    if (buffer.readUInt32LE(localOffset) !== LFH_SIG) throw new XlsxError(`zip 本地头损坏: ${name}`)
    const localNameLength = buffer.readUInt16LE(localOffset + 26)
    const localExtraLength = buffer.readUInt16LE(localOffset + 28)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    const raw = buffer.subarray(dataStart, dataStart + compressedSize)
    files.set(name, method === 8 ? inflateRawSync(raw) : Buffer.from(raw))
    offset += 46 + nameLength + extraLength + commentLength
  }
  return files
}

/** 解码 XML 实体。 */
function decodeEntities(text) {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, '&')
}

/**
 * 列字母转 0 基下标。
 * @param {string} letters 例如 A、B、AA
 * @returns {number}
 */
export function columnIndex(letters) {
  let value = 0
  for (const char of letters) value = value * 26 + (char.charCodeAt(0) - 64)
  return value - 1
}

/**
 * 解析 sharedStrings.xml。
 * @param {Buffer|undefined} xml
 * @returns {string[]}
 */
function parseSharedStrings(xml) {
  if (!xml) return []
  const text = xml.toString('utf8')
  const items = []
  for (const match of text.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const pieces = [...match[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((piece) => decodeEntities(piece[1]))
    items.push(pieces.join(''))
  }
  return items
}

/**
 * 解析工作表 XML 为二维数组。
 * @param {Buffer} xml
 * @param {string[]} shared
 * @returns {string[][]}
 */
function parseSheet(xml, shared) {
  const text = xml.toString('utf8')
  const rows = []
  for (const rowMatch of text.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
    const cells = []
    for (const cellMatch of rowMatch[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = cellMatch[1]
      const body = cellMatch[2] ?? ''
      const ref = /\br="([A-Z]+)\d+"/.exec(attrs)
      const type = /\bt="([^"]+)"/.exec(attrs)
      const target = ref ? columnIndex(ref[1]) : cells.length
      while (cells.length < target) cells.push('')
      let value = ''
      if (type && type[1] === 'inlineStr') {
        value = [...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((piece) => decodeEntities(piece[1])).join('')
      } else {
        const raw = /<v>([\s\S]*?)<\/v>/.exec(body)
        if (raw) value = type && type[1] === 's' ? (shared[Number(raw[1])] ?? '') : decodeEntities(raw[1])
      }
      cells[target] = value
    }
    rows.push(cells)
  }
  return rows
}

/**
 * 读取 xlsx 的第一个工作表。
 * @param {Buffer} buffer 文件内容
 * @returns {{headers: string[], rows: string[][]}}
 */
export function readXlsx(buffer) {
  const files = unzip(buffer)
  const sheetEntry = [...files.keys()].filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)).sort()[0]
  if (!sheetEntry) throw new XlsxError('xlsx 里没有工作表')
  const shared = parseSharedStrings(files.get('xl/sharedStrings.xml'))
  const rows = parseSheet(files.get(sheetEntry), shared)
  const nonEmpty = rows.filter((cells) => cells.some((cell) => String(cell).trim() !== ''))
  if (nonEmpty.length === 0) throw new XlsxError('工作表为空')
  const width = Math.max(...nonEmpty.map((cells) => cells.length))
  const padded = nonEmpty.map((cells) => {
    const copy = [...cells]
    while (copy.length < width) copy.push('')
    return copy
  })
  return { headers: padded[0].map((cell) => String(cell).trim()), rows: padded.slice(1) }
}
