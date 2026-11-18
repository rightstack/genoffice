/**
 * Speaker notes (notesSlide) read/write — archive surgery, same approach as
 * duplicateSlide.
 *
 * - Read: slide rels → notesSlide part → body placeholder <a:t> text (\n-separated).
 * - Write: patch the existing notesSlide's body txBody; if there is no notesSlide,
 *   create one (creating a notesMaster too if needed and registering it in
 *   presentation.xml).
 * All changes land in archive.entries: savePptx persists automatically, and the
 * main process's snapshot-style undo (shallow copy of entries) covers them.
 */
import type { OpenedPptx } from './index'
import { relsPathFor, resolveTarget, type PackageArchive } from './zip'
import {
  decodeNumericCharRefs,
  escapeXmlText,
  hasContentTypeOverride,
  maxRelationshipIdNumber,
} from './xml-utils'

const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NOTES_SLIDE_REL = `${REL_BASE}/notesSlide`
const NOTES_MASTER_REL = `${REL_BASE}/notesMaster`
const SLIDE_REL = `${REL_BASE}/slide`
const THEME_REL = `${REL_BASE}/theme`

const NOTES_SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const NOTES_MASTER_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'

function setEntry(archive: PackageArchive, path: string, xml: string): void {
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
}

/** Unescape XML text (for reading <a:t>). */
export function unescapeXml(s: string): string {
  return decodeNumericCharRefs(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** Path of the slide's notesSlide part (null if there is no notes slide). */
export function notesPathForSlide(archive: PackageArchive, slidePath: string): string | null {
  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type === NOTES_SLIDE_REL) return resolveTarget(slidePath, rel.target)
  }
  return null
}

/** Find the body placeholder sp block in the notesSlide xml. */
function findBodySp(xml: string): { xml: string; start: number; end: number } | null {
  for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    // quote-agnostic: a deck written with type='body' is still a body placeholder
    if (/<p:ph\b[^>]*\btype=["']body["']/.test(m[0])) {
      return { xml: m[0], start: m.index!, end: m.index! + m[0].length }
    }
  }
  return null
}

/** Read notes as plain text (body placeholder paragraphs joined by \n); '' if none. */
export function getSlideNotes(archive: PackageArchive, slidePath: string): string {
  const notesPath = notesPathForSlide(archive, slidePath)
  if (!notesPath) return ''
  const xml = archive.readText(notesPath)
  if (!xml) return ''
  const body = findBodySp(xml)
  if (!body) return ''
  const tx = /<p:txBody>([\s\S]*?)<\/p:txBody>/.exec(body.xml)?.[1]
  if (!tx) return ''
  const paras = [...tx.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) =>
    [...p[1]!.matchAll(/<a:t(?:\s[^>]*[^/>])?>([\s\S]*?)<\/a:t>/g)]
      .map((t) => unescapeXml(t[1]!))
      .join(''),
  )
  // Drop trailing empty paragraphs (PowerPoint templates often carry an empty placeholder paragraph)
  while (paras.length && paras[paras.length - 1] === '') paras.pop()
  return paras.join('\n')
}

/**
 * [lily] Notes as lightly formatted paragraphs: bold / italic runs, bullet or
 * auto-number lists and their level. Anything else in the XML (sizes, colours,
 * fields' semantics, line breaks) reads as plain text, like getSlideNotes.
 */
export interface NotesRun {
  text: string
  bold?: boolean
  italic?: boolean
}
export interface NotesParagraph {
  runs: NotesRun[]
  /** 'char' = <a:buChar>, 'num' = <a:buAutoNum>; absent = no bullet */
  bullet?: 'char' | 'num'
  /** 0-based list level (<a:pPr lvl>) */
  level?: number
}

const isOn = (attrs: string, name: string) =>
  new RegExp(`\\b${name}=["'](1|true)["']`).test(attrs)

/** [lily] Read notes as formatted paragraphs (same paragraphs as getSlideNotes); [] if none. */
export function getSlideNotesParagraphs(
  archive: PackageArchive,
  slidePath: string,
): NotesParagraph[] {
  const notesPath = notesPathForSlide(archive, slidePath)
  if (!notesPath) return []
  const xml = archive.readText(notesPath)
  if (!xml) return []
  const body = findBodySp(xml)
  if (!body) return []
  const tx = /<p:txBody>([\s\S]*?)<\/p:txBody>/.exec(body.xml)?.[1]
  if (!tx) return []
  const paras: NotesParagraph[] = [...tx.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) => {
    const inner = p[1]!
    const pPr = /<a:pPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:pPr>)/.exec(inner)
    const para: NotesParagraph = { runs: [] }
    if (pPr) {
      const lvl = Number(/\blvl=["'](\d+)["']/.exec(pPr[1]!)?.[1] ?? 0)
      const kids = pPr[2] ?? ''
      if (/<a:buAutoNum\b/.test(kids)) para.bullet = 'num'
      else if (/<a:buChar\b/.test(kids)) para.bullet = 'char'
      if (lvl > 0) para.level = lvl
    }
    for (const r of inner.matchAll(/<a:(r|fld)\b[^>]*>([\s\S]*?)<\/a:\1>/g)) {
      const text = [...r[2]!.matchAll(/<a:t(?:\s[^>]*[^/>])?>([\s\S]*?)<\/a:t>/g)]
        .map((t) => unescapeXml(t[1]!))
        .join('')
      if (!text) continue
      const rPr = /<a:rPr\b([^>]*?)\/?>/.exec(r[2]!)?.[1] ?? ''
      const run: NotesRun = { text }
      if (isOn(rPr, 'b')) run.bold = true
      if (isOn(rPr, 'i')) run.italic = true
      para.runs.push(run)
    }
    return para
  })
  while (paras.length && !paras[paras.length - 1]!.runs.length && !paras[paras.length - 1]!.bullet)
    paras.pop()
  return paras
}

/** [lily] One formatted paragraph → <a:p> (list indent: 0.25in per level, hanging bullet). */
function notesParagraphXml(p: NotesParagraph): string {
  const lvl = Math.max(0, Math.min(8, Math.floor(p.level ?? 0)))
  const pPr = p.bullet
    ? `<a:pPr marL="${228600 * (lvl + 1)}" lvl="${lvl}" indent="-228600">` +
      (p.bullet === 'num'
        ? '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>'
        : '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>') +
      '</a:pPr>'
    : ''
  const runs = p.runs
    .filter((r) => r.text !== '')
    .map(
      (r) =>
        `<a:r><a:rPr lang="zh-CN"${r.bold ? ' b="1"' : ''}${r.italic ? ' i="1"' : ''} dirty="0"/>` +
        `<a:t>${escapeXmlText(r.text)}</a:t></a:r>`,
    )
    .join('')
  return `<a:p>${pPr}${runs || '<a:endParaRPr lang="zh-CN"/>'}</a:p>`
}

function buildNotesTxBodyRich(paragraphs: NotesParagraph[]): string {
  const paras = paragraphs.length
    ? paragraphs.map(notesParagraphXml).join('')
    : '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody>`
}

/** text (\n-separated) → notes txBody. */
function buildNotesTxBody(text: string): string {
  const lines = text.split('\n')
  const paras = lines.every((l) => l === '')
    ? '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
    : lines
        .map((line) =>
          line === ''
            ? '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
            : `<a:p><a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${escapeXmlText(line)}</a:t></a:r></a:p>`,
        )
        .join('')
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody>`
}

const NOTES_BODY_SP_OPEN =
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder"/>' +
  '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'

/**
 * Write notes (overwrite): patch the existing notesSlide's body txBody;
 * if there is no notesSlide, create the part (with a notesMaster if needed).
 */
export function setSlideNotes(
  opened: OpenedPptx,
  slideIndex: number,
  text: string,
  /** [lily] formatted paragraphs; when given they are written instead of `text` */
  paragraphs?: NotesParagraph[],
): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  const { archive } = opened
  const notesPath = notesPathForSlide(archive, slide.path) ?? createNotesSlide(opened, slide.path)
  if (!notesPath) return false
  const xml = archive.readText(notesPath)
  if (!xml) return false

  const txBody = paragraphs ? buildNotesTxBodyRich(paragraphs) : buildNotesTxBody(text)
  const body = findBodySp(xml)
  let next: string
  if (body) {
    const patched = body.xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, () => txBody)
    next = xml.slice(0, body.start) + patched + xml.slice(body.end)
  } else {
    next = xml.replace('</p:spTree>', () => `${NOTES_BODY_SP_OPEN}${txBody}</p:sp></p:spTree>`)
  }
  setEntry(archive, notesPath, next)
  return true
}

/** Add an Override to [Content_Types].xml (skipped if already present). */
export function addContentTypeOverride(
  archive: PackageArchive,
  partPath: string,
  contentType: string,
): void {
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (!ct || hasContentTypeOverride(ct, partPath)) return
  setEntry(
    archive,
    ctPath,
    ct.replace(
      '</Types>',
      () => `<Override PartName="/${partPath}" ContentType="${contentType}"/></Types>`,
    ),
  )
}

/** Append a relationship to a part's rels, creating the rels file if missing. Returns the new rId. */
export function appendRelationship(
  archive: PackageArchive,
  partPath: string,
  type: string,
  target: string,
): string {
  const relsPath = relsPathFor(partPath)
  let xml =
    archive.readText(relsPath) ??
    XMLDECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  const maxRid = maxRelationshipIdNumber(xml)
  const rid = `rId${maxRid + 1}`
  xml = xml.replace(
    '</Relationships>',
    `<Relationship Id="${rid}" Type="${type}" Target="${target}"/></Relationships>`,
  )
  setEntry(archive, relsPath, xml)
  return rid
}

/** Ensure a notesMaster exists and return its part path. Created and registered in the presentation if missing. */
function ensureNotesMaster(archive: PackageArchive): string | null {
  for (const path of archive.entries.keys()) {
    if (/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) return path
  }
  const path = 'ppt/notesMasters/notesMaster1.xml'
  const emptyTree =
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
  const xml =
    XMLDECL +
    `<p:notesMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:cSld>${emptyTree}</p:cSld>` +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '</p:notesMaster>'
  setEntry(archive, path, xml)
  addContentTypeOverride(archive, path, NOTES_MASTER_CT)
  // rels: reuse an existing theme part (valid: multiple parts may point to the same theme)
  const theme = [...archive.entries.keys()].find((p) => /^ppt\/theme\/theme\d+\.xml$/.test(p))
  if (theme) appendRelationship(archive, path, THEME_REL, `../${theme.slice(4)}`)
  // Register notesMasterIdLst in presentation.xml
  const presPath = 'ppt/presentation.xml'
  const pres = archive.readText(presPath)
  if (pres) {
    const existingList = /<p:notesMasterIdLst\b[^>]*(?:\/>|>[\s\S]*?<\/p:notesMasterIdLst>)/.exec(
      pres,
    )?.[0]
    if (!existingList || existingList.endsWith('/>')) {
      const rid = appendRelationship(
        archive,
        presPath,
        NOTES_MASTER_REL,
        'notesMasters/notesMaster1.xml',
      )
      const list = `<p:notesMasterIdLst><p:notesMasterId r:id="${rid}"/></p:notesMasterIdLst>`
      const next = existingList
        ? pres.replace(existingList, () => list)
        : pres.includes('</p:sldMasterIdLst>')
          ? pres.replace('</p:sldMasterIdLst>', () => `</p:sldMasterIdLst>${list}`)
          : pres.replace('<p:sldIdLst>', () => `${list}<p:sldIdLst>`)
      setEntry(archive, presPath, next)
    }
  }
  return path
}

/** Create an empty notesSlide for a slide (part + rels + CT + slide rels); returns the part path. */
function createNotesSlide(opened: OpenedPptx, slidePath: string): string | null {
  const { archive } = opened
  if (!ensureNotesMaster(archive)) return null

  let maxNum = 0
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/notesSlides\/notesSlide(\d+)\.xml$/.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const notesPath = `ppt/notesSlides/notesSlide${maxNum + 1}.xml`

  const xml =
    XMLDECL +
    `<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    '<p:cSld><p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder"/>' +
    '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
    NOTES_BODY_SP_OPEN +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
  setEntry(archive, notesPath, xml)
  addContentTypeOverride(archive, notesPath, NOTES_SLIDE_CT)

  // notesSlide rels: notesMaster + owning slide
  const master = [...archive.entries.keys()].find((p) =>
    /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(p),
  )!
  appendRelationship(archive, notesPath, NOTES_MASTER_REL, `../${master.slice(4)}`)
  appendRelationship(archive, notesPath, SLIDE_REL, `../${slidePath.slice(4)}`)

  // slide rels: point to the new notesSlide
  appendRelationship(archive, slidePath, NOTES_SLIDE_REL, `../${notesPath.slice(4)}`)
  return notesPath
}
