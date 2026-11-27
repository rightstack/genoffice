/**
 * [lily] SmartArt previews whose sector shapes (pie / blockArc) have boxes far past the frame
 * — Target List quarter rings — keep their 1:1 canvas mapping: only what the sector draws counts
 * toward the stale-cache bounds, so the drawing is not compressed into the top-left corner.
 */
import { describe, it, expect } from 'vitest'
import type { PassthroughElement, Slide, SlideDeck } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'
import { sectorExtent } from '../src/build-slide'
import type { GroupRenderNode } from '../src/render-tree'

const size: SlideDeck['size'] = { cx: 9525 * 1280, cy: 9525 * 720 }
const anchor = { spIndex: 0, originalXml: '<p:graphicFrame/>', range: [0, 0] as [number, number] }
const xf = (x: number, y: number, cx: number, cy: number, rot = 0) => ({ offset: { x, y, cx, cy }, rot, flipH: false, flipV: false })
const slideOf = (elements: Slide['elements']): Slide => ({ path: 'ppt/slides/slide1.xml', originalXml: '', bodyPrefix: '', bodySuffix: '', elements })
const E = 9525
const pie = (id: string, x: number, y: number, d: number) => ({
  id,
  type: 'shape' as const,
  anchor: { spIndex: -1, originalXml: '', range: [0, 0] as [number, number] },
  transform: xf(x, y, d, d),
  presetGeometry: 'pie',
  // top → right quarter (270° → 0°)
  adjust: { adj1: 16200000, adj2: 0 },
  fill: { type: 'solid' as const, color: '#FFCC33' },
})

describe('sector-aware SmartArt preview bounds', () => {
  it('a top-right quarter pie reaches only its right edge and its centre line', () => {
    const r = sectorExtent(pie('p', 0, 0, 1000))!
    expect(r.right).toBeCloseTo(500, 0)
    expect(r.bottom).toBeCloseTo(0, 0)
    expect(sectorExtent({ ...pie('p', 0, 0, 1000), presetGeometry: 'rect' })).toBeNull()
    // turned 90° clockwise the same quarter now spans the bottom-right
    const t = sectorExtent({ ...pie('p', 0, 0, 1000), transform: xf(0, 0, 1000, 1000, 5400000) })!
    expect(t.right).toBeCloseTo(500, 0)
    expect(t.bottom).toBeCloseTo(500, 0)
  })

  it('quarter rings whose pie boxes overflow the frame keep the frame as the child space', () => {
    const fw = 600 * E
    const fh = 350 * E
    // rings centred on the frame's bottom-left corner: boxes reach left of and below the frame
    const el: PassthroughElement = {
      id: 'gf1',
      type: 'passthrough',
      kind: 'smartart',
      anchor,
      transform: xf(0, 0, fw, fh),
      previewShapes: [pie('a', -300 * E, 50 * E, 600 * E), pie('b', -200 * E, 150 * E, 400 * E), pie('c', -100 * E, 250 * E, 200 * E)],
    }
    const g = buildRenderSlide(slideOf([el]), size, { fitWidthPx: 1280 }).nodes[0] as GroupRenderNode
    // 1:1 — the largest pie keeps its 600 px box (it was squeezed to ~0.6× before)
    expect(g.children[0]!.box.w).toBeCloseTo(600, 0)
  })
})
