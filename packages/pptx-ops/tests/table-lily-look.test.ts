/**
 * [lily] Table look ops for the LilyOffice table panel: the `lilyGrid` preset (header row only, every
 * style option has a visible part), setTableStyle cellMargins / cellVert / clearTextFormatting.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  createBlankPptx,
  materializeSlide,
  openPptx,
  type OpenedPptx,
  type TableElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx

/** The table as reparsed from its current XML (setTableStyle patches the XML; the app reparses after a txn) */
const table = () => {
  materializeSlide(opened, 0)
  return opened.deck.slides[0]!.elements.find((el) => el.type === 'table') as TableElement
}
const xml = () => table().anchor.originalXml
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
})

describe('lilyGrid preset', () => {
  it('pins the Lily Table Grid style with header row only and draws the grid from the style', () => {
    expect(run({ op: 'setTableStyle', styleName: 'lilyGrid' }).applied).toBe(true)
    const t = table()
    expect(t.styleFlags).toMatchObject({ firstRow: true, bandRow: false, lastRow: false })
    expect(t.styleId).toBe('{A10FF1CE-0000-4000-9000-00000000011C}')
    const styles = String(opened.archive.entries.get('ppt/tableStyles.xml') ?? '')
    expect(styles).toContain('styleName="Lily Table Grid"')
    // header fill and bold, banded fill, grid lines all come from the style
    const cell = t.rows[0]![0]!
    expect(cell.fill).toMatchObject({ type: 'solid' })
  })

  it('lets every style option change the look (banded rows get the F5F5F5 band)', () => {
    run({ op: 'setTableStyle', styleName: 'lilyGrid' })
    expect(run({ op: 'setTableStyle', bandRow: true, lastRow: true }).applied).toBe(true)
    const t = table()
    expect(t.styleFlags).toMatchObject({ bandRow: true, lastRow: true })
    const fill = t.rows[1]!.map((c) => (c.fill as { color?: string } | undefined)?.color ?? null)
    expect(fill.every((c) => /F5F5F5/i.test(String(c)))).toBe(true)
  })
})

describe('setTableStyle cellMargins / cellVert', () => {
  it('writes tcPr insets on the given cells only (pt → EMU)', () => {
    expect(run({ op: 'setTableStyle', cellMargins: { t: 8, b: 8 }, cells: [{ row: 1, col: 1 }] }).applied).toBe(true)
    const t = table()
    expect(t.rows[1]![1]!.text?.insets).toMatchObject({ t: 101600, b: 101600, l: 91440 })
    expect(t.rows[0]![0]!.text?.insets).toMatchObject({ t: 45720 })
  })

  it('writes insets on every cell without cells and keeps other tcPr children', () => {
    run({ op: 'setTableStyle', shadingColor: '#FF0000' })
    expect(run({ op: 'setTableStyle', cellMargins: { l: 10 } }).applied).toBe(true)
    expect(table().rows.flat().every((c) => c.text?.insets?.l === 127000)).toBe(true)
    expect(xml()).toMatch(/<a:tcPr[^>]*marL="127000"[^>]*><a:solidFill>/)
  })

  it('rejects margins out of range', () => {
    expect(run({ op: 'setTableStyle', cellMargins: { l: -1 } }).applied).toBe(false)
    expect(run({ op: 'setTableStyle', cellMargins: {} }).applied).toBe(false)
  })

  it('sets and clears the cell text direction', () => {
    expect(run({ op: 'setTableStyle', cellVert: 'vert270', cells: [{ row: 0, col: 0 }] }).applied).toBe(true)
    expect(table().rows[0]![0]!.text?.vert).toBe('vert270')
    expect(run({ op: 'setTableStyle', cellVert: 'horz', cells: [{ row: 0, col: 0 }] }).applied).toBe(true)
    expect(table().rows[0]![0]!.text?.vert).toBeUndefined()
    expect(run({ op: 'setTableStyle', cellVert: 'sideways' }).applied).toBe(false)
  })
})

describe('setTableStyle clearTextFormatting', () => {
  it('drops bold, size and color from every run but keeps the text', () => {
    run({
      op: 'setTableCell',
      row: 0,
      col: 0,
      paragraphs: [{ runs: [{ text: 'Hi', bold: true, fontSize: 30, color: '#FF0000' }] }],
    })
    expect(xml()).toMatch(/b="1"/)
    expect(run({ op: 'setTableStyle', clearTextFormatting: true }).applied).toBe(true)
    const r = table().rows[0]![0]!.text!.paragraphs[0]!.runs[0]!
    expect(r.text).toBe('Hi')
    expect(xml()).not.toMatch(/<a:rPr[^>]*\bb="1"/)
    expect(xml()).not.toMatch(/<a:rPr[^>]*\bsz="/)
    expect(xml()).not.toMatch(/FF0000/)
  })
})
