/**
 * [lily] Table tools ops: tableMove (move a row/column to a boundary), setTableStyle cellSides
 * (per-edge cell borders, explicit no-line) and the cell-scoped `cells` on setFont / setParagraphFormat.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  createBlankPptx,
  openPptx,
  type OpenedPptx,
  type TableElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx

const table = () =>
  opened.deck.slides[0]!.elements.find((el) => el.type === 'table') as TableElement
const textOf = (r: number, c: number) =>
  (table().rows[r]![c]!.text?.paragraphs ?? [])
    .map((p) => p.runs.map((x) => x.text).join(''))
    .join('\n')
const run = (op: Record<string, unknown>) =>
  runTxn(opened, { ops: [{ target: { slide: 0, el: table().id }, ...op } as never] })

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
  const r = runTxn(opened, {
    ops: [
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 3,
        cols: 3,
        offset: { x: 0, y: 0, cx: 2743200, cy: 1371600 },
      },
    ],
  })
  expect(r.applied).toBe(true)
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      expect(
        run({ op: 'setTableCell', row, col, paragraphs: [{ runs: [{ text: `${row}${col}` }] }] })
          .applied,
      ).toBe(true)
    }
  }
})

describe('tableMove', () => {
  it('moves a row below the next one (boundary 2)', () => {
    expect(run({ op: 'tableMove', axis: 'row', from: 0, to: 2 }).applied).toBe(true)
    expect([textOf(0, 0), textOf(1, 0), textOf(2, 0)]).toEqual(['10', '00', '20'])
  })

  it('moves a column to the front, every row and the grid together', () => {
    const widths = () => table().anchor.originalXml.match(/<a:gridCol w="(\d+)"/g)
    run({ op: 'setTableColWidth', col: 2, wEmu: 1234567 })
    expect(run({ op: 'tableMove', axis: 'col', from: 2, to: 0 }).applied).toBe(true)
    expect([textOf(0, 0), textOf(0, 1), textOf(2, 0)]).toEqual(['02', '00', '22'])
    expect(widths()![0]).toContain('1234567')
  })

  it('refuses a no-op or out-of-range move', () => {
    expect(run({ op: 'tableMove', axis: 'row', from: 1, to: 1 }).applied).toBe(false)
    expect(run({ op: 'tableMove', axis: 'row', from: 1, to: 2 }).applied).toBe(false)
    expect(run({ op: 'tableMove', axis: 'col', from: 0, to: 9 }).applied).toBe(false)
    expect(run({ op: 'tableMove', axis: 'x', from: 0, to: 1 }).applied).toBe(false)
  })
})

describe('setTableStyle cellSides', () => {
  it('writes one edge with dash and compound, and an explicit no-line edge', () => {
    const r = run({
      op: 'setTableStyle',
      cellSides: [
        { row: 1, col: 1, side: 't', ln: { color: '#FF0000', widthPt: 2, dash: 'sysDash' } },
        { row: 1, col: 1, side: 'l', ln: { color: '#00FF00', widthPt: 1, compound: 'dbl' } },
        { row: 1, col: 1, side: 'b', ln: null },
      ],
    })
    expect(r.applied).toBe(true)
    const xml = table().anchor.originalXml
    expect(xml).toContain(
      '<a:lnT w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:lnT>',
    )
    expect(xml).toContain('<a:lnL w="12700" cmpd="dbl">')
    expect(xml).toContain('<a:lnB w="12700"><a:noFill/></a:lnB>')
    // schema order inside tcPr: lnL before lnT before lnB
    const tc = /<a:lnL w="12700" cmpd="dbl">.*?<\/a:tcPr>/s.exec(xml)![0]
    expect(tc.indexOf('<a:lnL')).toBeLessThan(tc.indexOf('<a:lnT'))
    expect(tc.indexOf('<a:lnT')).toBeLessThan(tc.indexOf('<a:lnB'))
  })

  it('replaces an existing edge instead of duplicating it', () => {
    run({
      op: 'setTableStyle',
      cellSides: [{ row: 0, col: 0, side: 'r', ln: { color: '#000000', widthPt: 1 } }],
    })
    run({
      op: 'setTableStyle',
      cellSides: [{ row: 0, col: 0, side: 'r', ln: { color: '#0000FF', widthPt: 3 } }],
    })
    const xml = table().anchor.originalXml
    expect(xml.match(/<a:lnR /g)).toHaveLength(1)
    expect(xml).toContain('<a:lnR w="38100"><a:solidFill><a:srgbClr val="0000FF"/>')
  })

  it('rejects a bad side or color', () => {
    expect(
      run({ op: 'setTableStyle', cellSides: [{ row: 0, col: 0, side: 'x', ln: null }] }).applied,
    ).toBe(false)
    expect(
      run({
        op: 'setTableStyle',
        cellSides: [{ row: 0, col: 0, side: 't', ln: { color: 'red', widthPt: 1 } }],
      }).applied,
    ).toBe(false)
  })
})

describe('setFont / setParagraphFormat cells', () => {
  it('formats only the listed cells', () => {
    expect(
      run({ op: 'setFont', font: { bold: true, fontSizePt: 20 }, cells: [{ row: 0, col: 1 }] })
        .applied,
    ).toBe(true)
    expect(table().rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    expect(table().rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.bold).toBeFalsy()
    expect(
      run({ op: 'setParagraphFormat', format: { align: 'center' }, cells: [{ row: 2, col: 2 }] })
        .applied,
    ).toBe(true)
    expect(table().rows[2]![2]!.text!.paragraphs[0]!.align).toBe('center')
    expect(table().rows[2]![1]!.text!.paragraphs[0]!.align).not.toBe('center')
  })

  it('rejects malformed cells', () => {
    expect(run({ op: 'setFont', font: { bold: true }, cells: [{ row: -1, col: 0 }] }).applied).toBe(
      false,
    )
  })
})
