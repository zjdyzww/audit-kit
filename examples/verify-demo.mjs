#!/usr/bin/env node
/**
 * @module demo_verify
 * @path ~/earn/audit-kit/examples/verify-demo.mjs
 * @description 生成演示结算单并独立核验审计器输出。
 *   关键：期望值由**解析式计算**（每笔 gmv/commission 固定），不是抄工具的输出——
 *   否则就是循环论证。
 * @zone verify
 * @tier Tier1-delivery
 *
 * @skill buildLines — 生成 117 笔正常单 + 3 笔异常单的 CSV 行
 * @skill expectations — 用解析式推出期望的 GMV/佣金/服务费/费率/异常计数
 * @skill verifyDemo — 跑审计器并与独立期望逐项比对，输出失败明细
 *
 * @rule independent_expectation P0 — 期望值必须由本文件的算式推出，禁止读取 audit 的输出反推
 *   @enforced-by verify-demo.mjs:自身断言
 * @rule anomaly_multiplicity P1 — 异常必须逐个计数比对，不只比"有异常"
 *   @enforced-by verify-demo.mjs:自身断言
 *
 * 用法：node examples/verify-demo.mjs
 * 退出码：0 全部符合 / 1 有偏差
 */

import { writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseCsv, mapRows, audit } from '../audit.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const HEADER = 'order_id,settle_date,gmv,commission,service_fee,refund,payout'

/** 正常单：交易额 1000，佣金 80，服务费 3，到账 917（费率 8.3%）。 */
const NORMAL = { gmv: 1000, commission: 80, serviceFee: 3, refund: 0, payout: 917 }
const NORMAL_COUNT = 117
/** 异常单三笔，各自触发不同异常码。 */
const BAD_PAYOUT = { id: 'BAD-PAYOUT', gmv: 1000, commission: 80, serviceFee: 3, refund: 0, payout: 999 }
const DUP_ID = 'N0001'
const BAD_COMMISSION = { id: 'BAD-COMMISSION', gmv: 1000, commission: 1200, serviceFee: 0, refund: 0, payout: -200 }

/** 生成 CSV 行。 */
function buildLines() {
  const lines = [HEADER]
  for (let index = 1; index <= NORMAL_COUNT; index += 1) {
    const id = `N${String(index).padStart(4, '0')}`
    const day = String((index % 28) + 1).padStart(2, '0')
    const month = String((index % 12) + 1).padStart(2, '0')
    lines.push(`${id},2026-${month}-${day},${NORMAL.gmv},${NORMAL.commission},${NORMAL.serviceFee},${NORMAL.refund},${NORMAL.payout}`)
  }
  lines.push(`${BAD_PAYOUT.id},2026-03-15,${BAD_PAYOUT.gmv},${BAD_PAYOUT.commission},${BAD_PAYOUT.serviceFee},${BAD_PAYOUT.refund},${BAD_PAYOUT.payout}`)
  lines.push(`${DUP_ID},2026-04-01,${NORMAL.gmv},${NORMAL.commission},${NORMAL.serviceFee},${NORMAL.refund},${NORMAL.payout}`)
  lines.push(`${BAD_COMMISSION.id},2026-05-20,${BAD_COMMISSION.gmv},${BAD_COMMISSION.commission},${BAD_COMMISSION.serviceFee},${BAD_COMMISSION.refund},${BAD_COMMISSION.payout}`)
  return lines
}

/** 解析式期望值（独立于工具实现）。 */
function expectations() {
  const rows = NORMAL_COUNT + 3
  const gmv = NORMAL_COUNT * NORMAL.gmv + BAD_PAYOUT.gmv + NORMAL.gmv + BAD_COMMISSION.gmv
  const commission = NORMAL_COUNT * NORMAL.commission + BAD_PAYOUT.commission + NORMAL.commission + BAD_COMMISSION.commission
  const serviceFee = NORMAL_COUNT * NORMAL.serviceFee + BAD_PAYOUT.serviceFee + NORMAL.serviceFee + BAD_COMMISSION.serviceFee
  const charges = commission + serviceFee
  return {
    rows,
    gmv,
    commission,
    serviceFee,
    charges,
    effectiveRate: charges / gmv,
    // 异常：BAD-PAYOUT → PAYOUT_MISMATCH；重复单 → DUPLICATE_ORDER；
    // BAD-COMMISSION → COMMISSION_GT_GMV + NEGATIVE_AMOUNT + RATE_OUTLIER
    anomalyCount: 5,
    expectedCodes: {
      PAYOUT_MISMATCH: 1,
      DUPLICATE_ORDER: 1,
      COMMISSION_GT_GMV: 1,
      NEGATIVE_AMOUNT: 1,
      RATE_OUTLIER: 1,
    },
  }
}

const lines = buildLines()
const csvPath = join(here, '演示-结算单.csv')
writeFileSync(csvPath, `${lines.join('\n')}\n`, 'utf8')

const expected = expectations()
const result = audit(mapRows(parseCsv(lines.join('\n'))), 0.06)
const failures = []
const near = (a, b) => Math.abs(a - b) < 1e-9

if (result.totals.gmv !== expected.gmv) failures.push(`GMV 期望 ${expected.gmv}，实际 ${result.totals.gmv}`)
if (result.totals.commission !== expected.commission) failures.push(`佣金 期望 ${expected.commission}，实际 ${result.totals.commission}`)
if (result.totals.serviceFee !== expected.serviceFee) failures.push(`服务费 期望 ${expected.serviceFee}，实际 ${result.totals.serviceFee}`)
if (!near(result.effectiveRate, expected.effectiveRate)) failures.push(`费率 期望 ${expected.effectiveRate}，实际 ${result.effectiveRate}`)
if (result.anomalies.length !== expected.anomalyCount) failures.push(`异常 期望 ${expected.anomalyCount} 条，实际 ${result.anomalies.length}`)
for (const [code, count] of Object.entries(expected.expectedCodes)) {
  const actual = result.anomalies.filter((item) => item.code === code).length
  if (actual !== count) failures.push(`异常码 ${code} 期望 ${count} 次，实际 ${actual}`)
}

console.log(`演示数据：${expected.rows} 笔，写入 ${csvPath}`)
console.log(`GMV ¥${expected.gmv.toLocaleString('zh-CN')}，抽成 ¥${expected.charges.toLocaleString('zh-CN')}，费率 ${(expected.effectiveRate * 100).toFixed(2)}%`)
console.log(`公示 6.00% 应抽 ¥${(expected.gmv * 0.06).toLocaleString('zh-CN')}，差额 ¥${(expected.charges - expected.gmv * 0.06).toLocaleString('zh-CN')}`)
if (failures.length > 0) {
  for (const item of failures) console.error(`DEMO FAIL: ${item}`)
  process.exit(1)
}
console.log('DEMO PASS：所有数值与异常计数均与独立算式一致')
