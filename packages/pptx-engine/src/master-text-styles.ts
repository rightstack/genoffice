/**
 * Slide-master text styles (<p:txStyles>): the title style and the first body levels
 * every placeholder inherits from. PowerPoint edits them through Slide Master view;
 * this writes size / bold straight into the master parts so a whole deck's title and
 * body text change at once, the way the master edit does.
 *
 * Entry surgery on the master XML (like theme-apply): only the touched attributes
 * change, unknown content (bullets, spacing, colors, extLst) is kept verbatim.
 */
import type { OpenedPptx } from './index'

export interface MasterLevelPatch {
  /** Font size (pt); written as sz in hundredths */
  sizePt?: number
  bold?: boolean
}

export interface MasterTextStyleSpec {
  /** titleStyle lvl1 */
  title?: MasterLevelPatch
  /** bodyStyle levels (1-based) */
  body?: Partial<Record<1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9, MasterLevelPatch>>
}

const MASTER_PART_RE = /^ppt\/slideMasters\/slideMaster\d+\.xml$/

/** Set sz / b on an rPr-like start tag (`<a:defRPr …>` or `<a:defRPr …/>`), keeping other attributes. */
function patchRPrTag(tag: string, p: MasterLevelPatch): string {
  let out = tag
  const setAttr = (name: string, value: string | null) => {
    const re = new RegExp(`\\s${name}="[^"]*"`)
    if (value == null) out = out.replace(re, '')
    else if (re.test(out)) out = out.replace(re, ` ${name}="${value}"`)
    else out = out.replace(/^<a:defRPr/, `<a:defRPr ${name}="${value}"`)
  }
  if (p.sizePt != null) setAttr('sz', String(Math.round(p.sizePt * 100)))
  if (p.bold != null) setAttr('b', p.bold ? '1' : '0')
  return out
}

/** Patch (or create) <a:defRPr> inside one <a:lvlNpPr> element. */
function patchLevelXml(lvlXml: string, n: number, p: MasterLevelPatch): string {
  // Self-closing level: open it up
  const selfClose = new RegExp(`^<a:lvl${n}pPr([^>]*)/>$`).exec(lvlXml)
  if (selfClose) lvlXml = `<a:lvl${n}pPr${selfClose[1]}></a:lvl${n}pPr>`
  const def = /<a:defRPr\b[^>]*?\/?>/.exec(lvlXml)
  if (def) return lvlXml.replace(def[0], patchRPrTag(def[0], p))
  // No defRPr: insert one before extLst, else before the closing tag (schema order: defRPr precedes extLst)
  const fresh = patchRPrTag('<a:defRPr/>', p)
  const ext = lvlXml.indexOf('<a:extLst')
  if (ext >= 0) return lvlXml.slice(0, ext) + fresh + lvlXml.slice(ext)
  return lvlXml.replace(new RegExp(`</a:lvl${n}pPr>$`), `${fresh}</a:lvl${n}pPr>`)
}

/** Patch level n inside a style section's inner XML (creating the level in order when missing). */
function patchSectionInner(inner: string, n: number, p: MasterLevelPatch): string {
  const re = new RegExp(`<a:lvl${n}pPr\\b[^>]*?(?:/>|>[\\s\\S]*?</a:lvl${n}pPr>)`)
  const m = re.exec(inner)
  if (m) return inner.replace(m[0], patchLevelXml(m[0], n, p))
  const created = patchLevelXml(`<a:lvl${n}pPr></a:lvl${n}pPr>`, n, p)
  // Keep lvl order: before the first higher level (or extLst), else at the end
  for (let k = n + 1; k <= 9; k++) {
    const at = inner.search(new RegExp(`<a:lvl${k}pPr\\b`))
    if (at >= 0) return inner.slice(0, at) + created + inner.slice(at)
  }
  const ext = inner.indexOf('<a:extLst')
  if (ext >= 0) return inner.slice(0, ext) + created + inner.slice(ext)
  return inner + created
}

function patchSection(
  xml: string,
  section: 'titleStyle' | 'bodyStyle',
  levels: Array<[number, MasterLevelPatch]>,
): string {
  if (!levels.length) return xml
  const tag = `p:${section}`
  let out = xml
  // <p:titleStyle/> → empty open/close pair
  out = out.replace(new RegExp(`<${tag}\\s*/>`), `<${tag}></${tag}>`)
  const re = new RegExp(`(<${tag}\\b[^>]*>)([\\s\\S]*?)(</${tag}>)`)
  const m = re.exec(out)
  if (!m) {
    // No section at all: add it inside <p:txStyles> (title before body)
    if (!/<p:txStyles\b/.test(out)) return out
    let inner = ''
    for (const [n, p] of levels) inner = patchSectionInner(inner, n, p)
    const created = `<${tag}>${inner}</${tag}>`
    if (section === 'titleStyle') return out.replace(/(<p:txStyles\b[^>]*>)/, `$1${created}`)
    return out.replace(/(<\/p:titleStyle>)/, `$1${created}`)
  }
  let inner = m[2]!
  for (const [n, p] of levels) inner = patchSectionInner(inner, n, p)
  return out.replace(m[0], `${m[1]}${inner}${m[3]}`)
}

/** Apply the spec to one master part's XML. */
export function patchMasterTextStylesXml(xml: string, spec: MasterTextStyleSpec): string {
  let out = xml
  if (spec.title) out = patchSection(out, 'titleStyle', [[1, spec.title]])
  if (spec.body) {
    const levels = Object.entries(spec.body)
      .map(([k, p]) => [Number(k), p] as [number, MasterLevelPatch])
      .filter(([n, p]) => n >= 1 && n <= 9 && p)
      .sort((a, b) => a[0] - b[0])
    out = patchSection(out, 'bodyStyle', levels)
  }
  return out
}

/** Apply the spec to every slide master in the package; returns the number of parts rewritten. */
export function applyMasterTextStylesToArchive(
  opened: OpenedPptx,
  spec: MasterTextStyleSpec,
): number {
  const { archive } = opened
  let patched = 0
  for (const path of [...archive.entries.keys()]) {
    if (!MASTER_PART_RE.test(path)) continue
    const xml = archive.readText(path)
    if (!xml) continue
    const next = patchMasterTextStylesXml(xml, spec)
    if (next !== xml) {
      archive.entries.set(path, Buffer.from(next, 'utf8'))
      patched++
    }
  }
  return patched
}
