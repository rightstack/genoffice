// [lily] lily-docs 차트와 같은 모양: 둥근 막대 끝 · 동그라미 범례 · 전체 불투명도 · 영역 midCat · 레이블 글자색 · 막대 대 원형 · 영역 그라데이션
import { describe, it, expect } from 'vitest'
import { parseChartXml, type ChartModel } from '@genoffice/pptx-engine'
import { buildChartNode, roundEndPath } from '../src/build-chart'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = { x: 100, y: 100, w: 600, h: 400, centerX: 400, centerY: 300, rotationDeg: 0, flipH: false, flipV: false }

describe('lily chart look', () => {
  it('rounded bars: only the value end of a clustered column is rounded, radius <= 4pt and <= a quarter of the thickness', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [40, -10] }],
      lily: { roundedBars: true },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.bars.map((b) => b.roundEnd)).toEqual(['top', 'bottom'])
    for (const b of node.bars) {
      expect(b.r!).toBeGreaterThan(0)
      expect(b.r!).toBeLessThanOrEqual(b.w * 0.25 + 1e-6)
    }
    expect(roundEndPath(node.bars[0]!)).toMatch(/ A /)
  })

  it('rounded bars: only the outermost segment of a stack is rounded; horizontal bars round the right end', () => {
    const stacked: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      categories: ['A'],
      series: [
        { name: 'a', values: [10] },
        { name: 'b', values: [20] },
      ],
      lily: { roundedBars: true },
    }
    const n1 = buildChartNode('r', 'e', stacked, box, vp, metrics)!
    expect(n1.bars.map((b) => b.roundEnd)).toEqual([undefined, 'top'])
    const horiz: ChartModel = { ...stacked, barDir: 'bar', grouping: 'clustered' }
    const n2 = buildChartNode('r', 'e', horiz, box, vp, metrics)!
    expect(n2.bars.every((b) => b.roundEnd === 'right')).toBe(true)
  })

  it('without the lily flags nothing is rounded and swatches stay square', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      categories: ['A'],
      series: [{ name: 'S', values: [40] }],
      legendPos: 't',
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.bars[0]!.roundEnd).toBeUndefined()
    expect(node.swatches.every((s) => !s.round)).toBe(true)
    expect(node.opacity).toBeUndefined()
  })

  it('legend circle + whole-chart opacity', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      categories: ['A'],
      series: [{ name: 'S', values: [40] }],
      legendPos: 't',
      lily: { legendCircle: true, opacity: 0.5 },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.swatches.length).toBeGreaterThan(0)
    expect(node.swatches.every((s) => s.round)).toBe(true)
    expect(node.opacity).toBe(0.5)
  })

  it('midCat areas span the plot from the left edge to the right edge', () => {
    const model: ChartModel = {
      kind: 'area',
      categories: ['1', '2', '3'],
      series: [{ name: 'S', values: [1, 3, 2] }],
      crossMidCat: true,
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const xs = node.polylines[0]!.points.filter((_, i) => i % 2 === 0)
    const between = buildChartNode('r', 'e', { ...model, crossMidCat: false }, box, vp, metrics)!
    const bx = between.polylines[0]!.points.filter((_, i) => i % 2 === 0)
    expect(Math.min(...xs)).toBeLessThan(Math.min(...bx))
    expect(Math.max(...xs)).toBeGreaterThan(Math.max(...bx))
    // the first point sits on the left axis line
    const axisX = Math.min(...node.axisLines.map((a) => a.x1))
    expect(Math.min(...xs)).toBeCloseTo(axisX, 0)
  })

  it('stacked data labels: white on a dark fill, dark on a light fill (lily only)', () => {
    const mk = (lily: boolean): ChartModel => ({
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      categories: ['A'],
      dataLabels: true,
      series: [
        { name: 'dark', values: [10], color: '#111111' },
        { name: 'light', values: [10], color: '#EEEEEE' },
      ],
      ...(lily ? { lily: {} } : {}),
    })
    const lilyNode = buildChartNode('r', 'e', mk(true), box, vp, metrics)!
    expect(lilyNode.labels.filter((l) => /^10$/.test(l.text)).map((l) => l.color).slice(-2)).toEqual(['#FFFFFF', '#1F1F1F'])
    const plain = buildChartNode('r', 'e', mk(false), box, vp, metrics)!
    expect(plain.labels.filter((l) => /^10$/.test(l.text)).map((l) => l.color).slice(-2)).toEqual(['#FFFFFF', '#FFFFFF'])
  })

  it('area fill: vertical gradient + a separate top-edge stroke', () => {
    const model: ChartModel = {
      kind: 'area',
      categories: ['1', '2'],
      series: [{ name: 'S', values: [1, 3], color: '#336699', fillAlpha: 0.5, fillGradAlpha: [0.8, 0.15], lineWidthPt: 1.5 }],
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const fillPoly = node.polylines.find((p) => p.closed)!
    expect(fillPoly.fillGradient?.stops).toHaveLength(2)
    expect(fillPoly.widthPx).toBe(0)
    expect(node.polylines.some((p) => !p.closed && p.widthPx > 0)).toBe(true)
  })

  it('bar of pie: the split points are drawn as a stacked bar, not a second pie', () => {
    const model: ChartModel = {
      kind: 'pie',
      categories: ['a', 'b', 'c', 'd', 'e'],
      series: [{ name: 'S', values: [30, 25, 20, 15, 10], pointColors: ['#111111', '#333333', '#555555', '#777777', '#999999', '#BBBBBB'] }],
      ofPie: { splitPos: 2, secondPieSize: 75, gapWidth: 100, bar: true },
      legendPos: 't',
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
    // d : e = 15 : 10 -> heights in that ratio, d on top
    expect(node.bars[0]!.h / node.bars[1]!.h).toBeCloseTo(1.5, 1)
    expect(node.bars[0]!.y).toBeLessThan(node.bars[1]!.y)
    const pie = buildChartNode('r', 'e', { ...model, ofPie: { ...model.ofPie!, bar: false } }, box, vp, metrics)!
    expect(pie.bars).toHaveLength(0)
  })

  it('bubbles: each bubble takes its own point colour and a see-through one gets a solid edge', () => {
    const model: ChartModel = {
      kind: 'scatter',
      categories: [],
      series: [
        {
          name: 'S',
          values: [1, 2],
          xValues: [1, 2],
          bubbleSizes: [5, 8],
          color: 'rgba(0, 0, 0, 0.6)',
          pointColors: ['rgba(255, 0, 0, 0.6)', 'rgba(0, 0, 255, 0.6)'],
        },
      ],
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.markers.map((m) => m.color)).toEqual(['rgba(255, 0, 0, 0.6)', 'rgba(0, 0, 255, 0.6)'])
    expect(node.markers.map((m) => m.stroke)).toEqual(['rgb(255, 0, 0)', 'rgb(0, 0, 255)'])
  })
})

describe('lily chart look: axis, chips, keys, radar', () => {
  const colModel = (lily: boolean): ChartModel => ({
    kind: 'bar',
    barDir: 'col',
    categories: ['A', 'B'],
    series: [{ name: 'S', values: [130, 150] }],
    ...(lily ? { lily: {} } : {}),
  })
  const tickTexts = (n: { labels: Array<{ text: string }> }) => n.labels.map((l) => l.text).filter((t) => /^\d+$/.test(t))

  it('value axis: lily charts use the docs rule (150 -> 175), others keep the PowerPoint rule (160)', () => {
    expect(tickTexts(buildChartNode('r', 'e', colModel(true), box, vp, metrics)!)).toContain('175')
    const plain = tickTexts(buildChartNode('r', 'e', colModel(false), box, vp, metrics)!)
    expect(plain).toContain('160')
    expect(plain).not.toContain('175')
  })

  it('bubble axis: lily adds no bubble headroom', () => {
    const mk = (lily: boolean): ChartModel => ({
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [30, 88], xValues: [14, 46], bubbleSizes: [5, 8] }],
      ...(lily ? { lily: {} } : {}),
    })
    const ys = (n: { labels: Array<{ text: string }> }) => Math.max(...tickTexts(n).map(Number))
    expect(ys(buildChartNode('r', 'e', mk(true), box, vp, metrics)!)).toBe(100)
    expect(ys(buildChartNode('r', 'e', mk(false), box, vp, metrics)!)).toBeGreaterThan(100)
  })

  it('scatter x axis (docs): starts 5% of the span below the first point rounded to the step; narrow line ranges do not reach zero', () => {
    const sc: ChartModel = {
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [30, 88], xValues: [14, 46], bubbleSizes: [5, 8] }],
      lily: {},
    }
    const xs = buildChartNode('r', 'e', sc, box, vp, metrics)!.labels.map((l) => l.text).filter((t) => /^\d+$/.test(t)).map(Number)
    // y ticks come first (0..100), then the x ticks: 10, 15 … 50
    expect(xs.slice(-3)).toEqual([40, 45, 50])
    expect(xs.slice(-9)[0]).toBe(10)
    const ln: ChartModel = { kind: 'line', categories: ['a', 'b', 'c'], series: [{ name: 'S', values: [1000, 1020, 1050] }], lily: {} }
    const ys = tickTexts(buildChartNode('r', 'e', ln, box, vp, metrics)!).map(Number)
    expect(Math.min(...ys)).toBeGreaterThan(0)
  })

  it('waterfall: a short floating bar puts its label past the end in dark text; a tall one keeps it inside', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      categories: ['a', 'b'],
      dataLabels: true,
      series: [
        { name: 'base', values: [0, 100], noFill: true },
        { name: 'v', values: [300, 2], color: '#111111' },
      ],
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const tall = node.labels.filter((l) => l.text === '300').slice(-1)[0]!
    const short = node.labels.filter((l) => l.text === '2').slice(-1)[0]!
    expect(tall.color).toBe('#FFFFFF')
    expect(short.color).not.toBe('#FFFFFF')
    const shortBar = node.bars.find((b) => b.h < 10)!
    expect(short.y + short.fontSizePx).toBeLessThanOrEqual(shortBar.y + 0.5)
  })

  it('combo: a line label on a column is a white chip in the series colour; off a column it is plain', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      categories: ['A', 'B'],
      dataLabels: true,
      series: [
        { name: 'bars', values: [100, 5], color: '#111111' },
        { name: 'line', values: [50, 90], plotKind: 'line', color: '#CC3300' },
      ],
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const chips = node.labels.filter((l) => l.round)
    expect(chips.length).toBeGreaterThan(0)
    expect(chips.every((l) => l.fill === '#FFFFFF' && l.stroke === '#CC3300' && l.color === '#1F1F1F')).toBe(true)
    expect(node.labels.every((l) => !l.chipOf)).toBe(true)
    expect(node.labels.filter((l) => !l.round && /^(50|90)$/.test(l.text)).length + chips.length).toBe(2)
  })

  it('legend: line series keys are circles when the legend shape is circle', () => {
    const model: ChartModel = {
      kind: 'line',
      categories: ['A', 'B'],
      series: [{ name: 'L', values: [1, 2] }],
      legendPos: 't',
      lily: { legendCircle: true },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.swatches.length).toBe(1)
    expect(node.swatches[0]!.round).toBe(true)
  })

  it('filled radar: a circular gradient fades from the centre out', () => {
    const model: ChartModel = {
      kind: 'radar',
      radarStyle: 'filled',
      categories: ['a', 'b', 'c'],
      series: [{ name: 'S', values: [1, 2, 3], color: '#336699', fillRadialAlpha: [0.08, 0.35] }],
      lily: {},
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const poly = node.polylines.find((p) => p.fillRadial)!
    expect(poly.fillRadial!.stops[0]!.color).toContain('0.08')
    expect(poly.fillRadial!.stops[1]!.color).toContain('0.35')
  })
})

describe('lily chart look: parsed from the chart part', () => {
  const part = (spec: object, extra = '') =>
    `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:areaChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="336699"><a:alpha val="80000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="336699"><a:alpha val="15000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:ln w="19050"><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln></c:spPr><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:areaChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/>${extra}<c:crossBetween val="midCat"/></c:valAx></c:plotArea></c:chart><c:extLst><c:ext uri="{7B1E5C3A-4D2F-4E8A-9C61-1F0A2B3C4D5E}"><lily:spec xmlns:lily="urn:lilyoffice:chart">${JSON.stringify(spec).replace(/"/g, '&quot;')}</lily:spec></c:ext></c:extLst></c:chartSpace>`

  it('reads the look flags from the extension, midCat from the value axis, and the area gradient', () => {
    const m = parseChartXml(part({ roundedBars: true, legendShape: 'circle', opacity: 0.4 }))!
    expect(m.lily).toEqual({ roundedBars: true, legendCircle: true, opacity: 0.4 })
    expect(m.crossMidCat).toBe(true)
    expect(m.series[0]!.fillGradAlpha).toEqual([0.8, 0.15])
  })

  it('no extension: no lily look', () => {
    const xml = part({}).replace(/<c:extLst>.*<\/c:extLst>/, '')
    expect(parseChartXml(xml)!.lily).toBeUndefined()
  })
})
