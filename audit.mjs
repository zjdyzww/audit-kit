#!/usr/bin/env node
/**
 * @module audit_kit
 * @path ~/earn/audit-kit/audit.mjs
 * @description 抽成审计器 v0：读平台结算单 CSV，算实际抽成费率、与公示费率比对、列出异常。
 *   纯本地运行，不联网、不上传、不留存。规格见同目录 spec.md。
 * @zone action
 * @tier Tier1-delivery
 *
 * @skill parseCsv — 解析 CSV（容忍 BOM、引号、空行），返回 {headers, rows}
 * @skill mapRows — 按映射把平台列名映射到标准账目模型；缺列即抛 InputError
 * @skill suggestMapping — 从实际表头猜一份映射草案（只猜名字，不改数据）
 * @skill audit — 汇总 GMV/佣金/服务费/退款/到账，算有效费率、应抽额、差额、异常
 * @skill runSelfTest — 正例 + 六条负例（含映射错误），真实退出码
 *
 * @rule missing_field_is_fatal P0 — 必需列缺失（映射后仍缺）⇒ InputError，exit 2，不产出半份报告
 *   @enforced-by audit.mjs:--self-test
 * @rule mapping_must_resolve P0 — --mapping 指向的列名必须在 CSV 里存在，否则 exit 2 并列出实际列名
 *   @enforced-by audit.mjs:--self-test
 * @rule must_exist_and_finite P0 — 金额必须是有限数且非负；解析失败指出行列与原始值 ⇒ exit 2
 *   @enforced-by audit.mjs:--self-test
 * @rule anomaly_nonblocking P1 — 异常不中止汇总，计入 anomalies 并让退出码为 1
 *   @enforced-by audit.mjs:--self-test
 * @rule payout_reconciliation P0 — payout 与 gmv-佣金-服务费+退款 的差 >0.01 即 PAYOUT_MISMATCH
 *   @enforced-by audit.mjs:--self-test
 * @rule no_silent_adapter_guess P0 — 适配器缺失或列对不上时列出实际列名，不猜、不填 0
 *   @enforced-by audit.mjs:--self-test
 *
 * 用法：
 *   node audit.mjs --input settle.csv [--declared-rate 0.06] [--adapter generic] [--json out.json]
 *   node audit.mjs --self-test
 */

import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { readXlsx, XlsxError } from './xlsx-lite.mjs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/** 标准账目模型的必需列。 */
const REQUIRED = ['order_id', 'settle_date', 'gmv', 'commission', 'service_fee', 'refund', 'payout']

/** v0 只内置通用适配器；平台适配器在拿到真实样本前不写死。 */
const ADAPTERS = {
  generic: null,
}

/** 输入不满足前置条件（exit 2）。 */
class InputError extends Error {}

/**
 * 解析 CSV 文本。
 * @param {string} text 原始文本（可含 BOM）
 * @returns {{headers: string[], rows: string[][]}}
 */
function parseCsv(text) {
  const clean = text.replace(/^\uFEFF/, '')
  const rows = []
  let row = []
  let field = ''
  let inQuotes = false
  for (let index = 0; index < clean.length; index += 1) {
    const ch = clean[index]
    if (inQuotes) {
      if (ch === '"' && clean[index + 1] === '"') { field += '"'; index += 1 }
      else if (ch === '"') inQuotes = false
      else field += ch
    } else if (ch === '"') inQuotes = true
    else if (ch === ',') { row.push(field); field = '' }
    else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = '' }
    else if (ch !== '\r') field += ch
  }
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row) }
  const nonEmpty = rows.filter((cells) => cells.some((cell) => String(cell).trim() !== ''))
  if (nonEmpty.length === 0) throw new InputError('CSV 为空')
  const headers = nonEmpty[0].map((cell) => String(cell).trim())
  return { headers, rows: nonEmpty.slice(1) }
}

/**
 * 把日期归一成 YYYY-MM-DD；xlsx 里的日期是 Excel 序列号，需换算。
 * 序列号合法区间取 20000–80000（约 1954–2119 年），区间外按原样保留。
 * @param {string} text 原始值
 * @param {number} lineNumber 行号（用于报错定位）
 * @returns {string}
 */
function normalizeDate(text, lineNumber) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text
  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text)
    if (serial >= 20000 && serial <= 80000) {
      const millis = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000
      return new Date(millis).toISOString().slice(0, 10)
    }
  }
  if (/^\d{4}[/.]\d{1,2}[/.]\d{1,2}$/.test(text)) {
    const [year, month, day] = text.split(/[/.]/)
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
  }
  throw new InputError(`第 ${lineNumber} 行 settle_date 无法解析: ${JSON.stringify(text)}`)
}

/**
 * 把原始行映射成标准账目记录。
 * @param {{headers: string[], rows: string[][]}} table
 * @param {Record<string,string>|null} mapping 标准字段 → 平台列名；null 表示列名已是标准名
 * @returns {Array<Record<string, string|number>>}
 */
function mapRows(table, mapping = null) {
  const present = new Set(table.headers)
  const resolve = (name) => (mapping && mapping[name] ? mapping[name] : name)
  const missing = REQUIRED.filter((name) => !present.has(resolve(name)))
  if (missing.length > 0) {
    const detail = missing.map((name) => (mapping && mapping[name] ? `${name}←「${mapping[name]}」` : name))
    throw new InputError(`缺少必需列: ${detail.join(', ')}；实际列: ${table.headers.join(', ')}`)
  }
  const position = Object.fromEntries(REQUIRED.map((name) => [name, table.headers.indexOf(resolve(name))]))
  return table.rows.map((cells, rowIndex) => {
    const record = {}
    for (const name of REQUIRED) {
      const raw = cells[position[name]] ?? ''
      if (name === 'order_id' || name === 'settle_date') {
        if (String(raw).trim() === '') throw new InputError(`第 ${rowIndex + 2} 行 ${name} 为空`)
        const text = String(raw).trim()
        record[name] = name === 'settle_date' ? normalizeDate(text, rowIndex + 2) : text
      } else {
        const value = Number(String(raw).replace(/,/g, '').trim())
        if (!Number.isFinite(value)) {
          throw new InputError(`第 ${rowIndex + 2} 行 ${name} 不是有效数字: ${JSON.stringify(raw)}`)
        }
        record[name] = value
      }
    }
    return record
  })
}

/**
 * 汇总并检查异常。
 * @param {Array<Record<string, string|number>>} records
 * @param {number|null} declaredRate 公示费率（小数），null 表示未提供
 * @returns {{totals: object, effectiveRate: number, expectedCommission: number|null, delta: number|null, anomalies: object[]}}
 */
function audit(records, declaredRate) {
  const totals = { gmv: 0, commission: 0, serviceFee: 0, refund: 0, payout: 0 }
  const anomalies = []
  const seen = new Set()
  for (const record of records) {
    const gmv = Number(record.gmv)
    const commission = Number(record.commission)
    const serviceFee = Number(record.service_fee)
    const refund = Number(record.refund)
    const payout = Number(record.payout)
    totals.gmv += gmv
    totals.commission += commission
    totals.serviceFee += serviceFee
    totals.refund += refund
    totals.payout += payout
    const push = (code, detail) => anomalies.push({ code, orderId: record.order_id, detail })
    if ([gmv, commission, serviceFee, refund, payout].some((value) => value < 0)) push('NEGATIVE_AMOUNT', '存在负金额')
    if (commission > gmv) push('COMMISSION_GT_GMV', `佣金 ${commission} > 交易额 ${gmv}`)
    if (refund > gmv) push('REFUND_GT_GMV', `退款 ${refund} > 交易额 ${gmv}`)
    const expected = gmv - commission - serviceFee + refund
    if (Math.abs(payout - expected) > 0.01) push('PAYOUT_MISMATCH', `到账 ${payout} ≠ 应有 ${expected.toFixed(2)}`)
    if (gmv > 0 && (commission + serviceFee) / gmv > 0.3) push('RATE_OUTLIER', `单笔费率 ${(((commission + serviceFee) / gmv) * 100).toFixed(1)}%`)
    if (seen.has(record.order_id)) push('DUPLICATE_ORDER', '订单号重复')
    seen.add(record.order_id)
  }
  const charges = totals.commission + totals.serviceFee
  const effectiveRate = totals.gmv > 0 ? charges / totals.gmv : 0
  const expectedCommission = declaredRate === null ? null : totals.gmv * declaredRate
  const delta = expectedCommission === null ? null : charges - expectedCommission
  return { totals, effectiveRate, expectedCommission, delta, anomalies }
}

/** 格式化金额。 */
function money(value) {
  return `¥${value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
}

/**
 * 从实际表头猜一份映射草案。只做名字匹配，不读数据、不猜测语义之外的变换。
 * @param {string[]} headers CSV 表头
 * @returns {{mapping: Record<string,string>, unresolved: string[]}}
 */
function suggestMapping(headers) {
  const aliases = {
    order_id: ['order_id', '订单号', '订单编号', '订单id', '交易单号', '子订单号'],
    settle_date: ['settle_date', '结算日期', '结算时间', '入账日期', '账期'],
    gmv: ['gmv', '交易额', '商品金额', '订单金额', '营业额', '成交金额'],
    commission: ['commission', '佣金', '平台服务费', '技术服务费', '抽成'],
    service_fee: ['service_fee', '服务费', '其他费用', '推广费', '履约服务费'],
    refund: ['refund', '退款', '退款金额', '退货金额'],
    payout: ['payout', '到账', '实收', '结算金额', '应结金额', '实际到账'],
  }
  const normalized = headers.map((header) => String(header).replace(/\s/g, '').toLowerCase())
  const mapping = {}
  const unresolved = []
  for (const [field, candidates] of Object.entries(aliases)) {
    const hit = candidates.find((candidate) => normalized.includes(candidate.toLowerCase()))
    if (hit) mapping[field] = headers[normalized.indexOf(hit.toLowerCase())]
    else if (normalized.includes(field)) mapping[field] = headers[normalized.indexOf(field)]
    else unresolved.push(field)
  }
  return { mapping, unresolved }
}

/**
 * 自检：正例 + 六条负例。
 * @returns {number} 0 通过 / 3 失败
 */
function runSelfTest() {
  const failures = []
  const header = REQUIRED.join(',')
  const good = [
    header,
    'A1,2026-01-05,1000,50,10,0,940',
    'A2,2026-01-06,2000,100,20,0,1880',
    'A3,2026-01-07,500,25,5,100,570',
  ].join('\n')
  const table = parseCsv(good)
  const records = mapRows(table)
  const result = audit(records, 0.06)
  if (result.totals.gmv !== 3500) failures.push(`GMV 应为 3500，实际 ${result.totals.gmv}`)
  if (Math.abs(result.effectiveRate - 210 / 3500) > 1e-9) failures.push(`费率应为 ${210 / 3500}，实际 ${result.effectiveRate}`)
  if (Math.abs(result.delta - (210 - 3500 * 0.06)) > 1e-9) failures.push(`差额算错: ${result.delta}`)
  if (result.anomalies.length !== 0) failures.push(`正例不应有异常，实际 ${result.anomalies.length} 条`)

  // 负例 1：缺列 ⇒ InputError
  try {
    mapRows(parseCsv('order_id,gmv\nA1,100'))
    failures.push('负例1 失败：缺列却未报错')
  } catch (error) {
    if (!(error instanceof InputError)) failures.push('负例1 抛出的不是 InputError')
  }
  // 负例 2：佣金 > GMV ⇒ 异常 COMMISSION_GT_GMV
  const bad1 = audit(mapRows(parseCsv(`${header}\nA1,2026-01-05,100,200,0,0,0`)), null)
  if (!bad1.anomalies.some((item) => item.code === 'COMMISSION_GT_GMV')) failures.push('负例2 失败：未检出 佣金>GMV')
  // 负例 3：到账对不上 ⇒ PAYOUT_MISMATCH
  const bad2 = audit(mapRows(parseCsv(`${header}\nA1,2026-01-05,1000,50,10,0,999`)), null)
  if (!bad2.anomalies.some((item) => item.code === 'PAYOUT_MISMATCH')) failures.push('负例3 失败：未检出 到账不符')
  // 负例 4：金额非数字 ⇒ InputError
  try {
    mapRows(parseCsv(`${header}\nA1,2026-01-05,abc,50,10,0,940`))
    failures.push('负例4 失败：非数字金额被接受')
  } catch (error) {
    if (!(error instanceof InputError)) failures.push('负例4 抛出的不是 InputError')
  }
  // 负例 5：重复订单号
  const bad3 = audit(mapRows(parseCsv(`${header}\nA1,2026-01-05,1000,50,10,0,940\nA1,2026-01-06,1000,50,10,0,940`)), null)
  if (!bad3.anomalies.some((item) => item.code === 'DUPLICATE_ORDER')) failures.push('负例5 失败：未检出 重复订单号')

  // 正例 2：列名是中文，靠映射跑通
  const zhHeader = '订单号,结算日期,商品金额,佣金,服务费,退款,实际到账'
  const zhMapping = { order_id: '订单号', settle_date: '结算日期', gmv: '商品金额', commission: '佣金', service_fee: '服务费', refund: '退款', payout: '实际到账' }
  const zhResult = audit(mapRows(parseCsv(`${zhHeader}\nB1,2026-02-01,1000,60,0,0,940`), zhMapping), 0.06)
  if (zhResult.totals.gmv !== 1000) failures.push(`映射正例 GMV 应为 1000，实际 ${zhResult.totals.gmv}`)
  if (Math.abs(zhResult.effectiveRate - 0.06) > 1e-9) failures.push(`映射正例 费率应为 0.06，实际 ${zhResult.effectiveRate}`)

  // 负例 6：映射指向不存在的列 ⇒ InputError
  try {
    mapRows(parseCsv(`${header}\nA1,2026-01-05,1000,50,10,0,940`), { gmv: '不存在的列' })
    failures.push('负例6 失败：映射指向不存在列却未报错')
  } catch (error) {
    if (!(error instanceof InputError)) failures.push('负例6 抛出的不是 InputError')
  }

  // mapping 草案：中文表头应能认出大部分字段
  const suggestion = suggestMapping(zhHeader.split(','))
  if (suggestion.unresolved.length > 0) failures.push(`映射草案未能识别: ${suggestion.unresolved.join(', ')}`)
  if (Object.keys(suggestion.mapping).length !== REQUIRED.length) failures.push(`映射草案应覆盖 ${REQUIRED.length} 个字段，实际 ${Object.keys(suggestion.mapping).length}`)

  // xlsx 正例：读 openpyxl 生成的 fixture（跨实现验证，不是自己读自己）
  try {
    const fixture = new URL('./examples/示例-结算单.xlsx', import.meta.url)
    const table = readXlsx(readFileSync(fixture))
    const mapping = suggestMapping(table.headers)
    if (mapping.unresolved.length > 0) failures.push(`xlsx 映射草案未识别: ${mapping.unresolved.join(', ')}`)
    const xlsxResult = audit(mapRows(table, mapping.mapping), 0.06)
    if (xlsxResult.totals.gmv !== 3500) failures.push(`xlsx 正例 GMV 应为 3500，实际 ${xlsxResult.totals.gmv}`)
    if (Math.abs(xlsxResult.effectiveRate - 290.5 / 3500) > 1e-9) failures.push(`xlsx 正例费率算错: ${xlsxResult.effectiveRate}`)
    const dates = mapRows(table, mapping.mapping).map((row) => row.settle_date)
    if (dates[0] !== '2026-01-05') failures.push(`xlsx 日期未从序列号归一: ${dates[0]}`)
  } catch (error) {
    failures.push(`xlsx 正例抛异常: ${error.message}`)
  }

  // xlsx 负例：不是 zip 的文件必须抛 XlsxError（不能返回空表）
  try {
    readXlsx(Buffer.from('这不是一个 xlsx'))
    failures.push('xlsx 负例失败：坏文件被接受')
  } catch (error) {
    if (!(error instanceof XlsxError)) failures.push('xlsx 负例抛出的不是 XlsxError')
  }

  if (failures.length > 0) {
    for (const item of failures) console.error(`SELFTEST FAIL: ${item}`)
    return 3
  }
  console.log('SELFTEST PASS: 正例(标准列/中文映射/真实 xlsx)全部正确；六条负例(缺列/非数字/佣金超额/到账不符/重复单号/映射列不存在)与 xlsx 坏文件负例均被拦下')
  return 0
}

/**
 * CLI 入口。
 * @param {string[]} argv
 * @returns {number} 退出码 0/1/2/3
 */
function main(argv) {
  if (argv.includes('--self-test')) {
    // 自检内部抛出的任何异常都必须归约为 exit 3，不能崩成未捕获异常（否则退出码语义失真）
    try {
      return runSelfTest()
    } catch (error) {
      console.error(`SELFTEST FAIL: 自检过程抛出异常 — ${error.message}`)
      return 3
    }
  }
  const option = (name, fallback = null) => {
    const index = argv.indexOf(`--${name}`)
    return index >= 0 && argv[index + 1] ? argv[index + 1] : fallback
  }
  const input = option('input')
  if (!input) {
    console.error('用法: node audit.mjs --input settle.csv [--declared-rate 0.06] [--json out.json] | --self-test')
    return 2
  }
  const adapter = option('adapter', 'generic')
  if (!(adapter in ADAPTERS)) {
    console.error(`适配器 ${adapter} 不在 v0 支持范围（仅 generic）。未拿到真实导出样本前不猜测列名。`)
    return 2
  }
  const rateText = option('declared-rate')
  const declaredRate = rateText === null ? null : Number(rateText)
  if (declaredRate !== null && !Number.isFinite(declaredRate)) {
    console.error(`--declared-rate 不是有效数字: ${rateText}`)
    return 2
  }
  try {
    const raw = readFileSync(input)
    const isXlsx = /\.xlsx$/i.test(input)
    const table = isXlsx ? readXlsx(raw) : parseCsv(raw.toString('utf8'))
    if (argv.includes('--emit-mapping')) {
      const { mapping, unresolved } = suggestMapping(table.headers)
      console.log(JSON.stringify(mapping, null, 2))
      if (unresolved.length > 0) {
        console.error(`未能自动识别: ${unresolved.join(', ')}；请手工补全后再用 --mapping 传入`)
        console.error(`实际列: ${table.headers.join(', ')}`)
        return 1
      }
      return 0
    }
    const mappingPath = option('mapping')
    let mapping = null
    if (mappingPath) {
      try {
        mapping = JSON.parse(readFileSync(mappingPath, 'utf8'))
      } catch (error) {
        console.error(`映射文件无法解析: ${mappingPath} (${error.message})`)
        return 2
      }
    }
    const result = audit(mapRows(table, mapping), declaredRate)
    const { totals, effectiveRate, expectedCommission, delta, anomalies } = result
    console.log(`实际抽成      ${money(totals.commission + totals.serviceFee)}   费率 ${(effectiveRate * 100).toFixed(2)}%`)
    console.log(`GMV           ${money(totals.gmv)}`)
    if (expectedCommission !== null) {
      console.log(`公示费率      ${(declaredRate * 100).toFixed(2)}%   应抽 ${money(expectedCommission)}`)
      console.log(`差额          ${money(delta)}`)
    } else {
      console.log('公示费率      未提供（加 --declared-rate 可算差额）')
    }
    console.log(`异常          ${anomalies.length} 项`)
    for (const item of anomalies.slice(0, 10)) console.log(`  - ${item.code} ${item.orderId} ${item.detail}`)
    const jsonPath = option('json')
    if (jsonPath) {
      writeFileSync(jsonPath, JSON.stringify({ totals, effectiveRate, declaredRate, expectedCommission, delta, anomalies }, null, 2), 'utf8')
      console.log(`JSON          ${jsonPath}`)
    }
    return anomalies.length > 0 ? 1 : 0
  } catch (error) {
    if (error instanceof InputError || error instanceof XlsxError) {
      console.error(`输入错误: ${error.message}`)
      return 2
    }
    throw error
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  process.exit(main(process.argv.slice(2)))
}

export { parseCsv, mapRows, audit, runSelfTest, InputError, REQUIRED }
