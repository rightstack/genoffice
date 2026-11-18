/**
 * [lily] Formatted speaker notes: bold / italic runs and bullet / number lists
 * survive write → read and save → reopen; plain text stays readable by getSlideNotes.
 */
import { describe, it, expect } from 'vitest'
import {
  createBlankPptx,
  getSlideNotes,
  getSlideNotesParagraphs,
  openPptx,
  savePptx,
  setSlideNotes,
  type NotesParagraph,
} from '../src/index'

const PARAS: NotesParagraph[] = [
  { runs: [{ text: 'Intro ' }, { text: 'bold', bold: true }, { text: ' and ' }, { text: 'it', italic: true }] },
  { runs: [{ text: 'first' }], bullet: 'char' },
  { runs: [{ text: 'nested', bold: true, italic: true }], bullet: 'char', level: 1 },
  { runs: [{ text: 'step <1> & "2"' }], bullet: 'num' },
  { runs: [] },
  { runs: [{ text: 'after gap' }] },
]

describe('notes paragraphs', () => {
  it('reads back what setSlideNotes wrote, and keeps the plain-text view', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setSlideNotes(opened, 0, 'ignored', PARAS)).toBe(true)
    const path = opened.deck.slides[0]!.path
    expect(getSlideNotesParagraphs(opened.archive, path)).toEqual(PARAS)
    expect(getSlideNotes(opened.archive, path)).toBe(
      'Intro bold and it\nfirst\nnested\nstep <1> & "2"\n\nafter gap',
    )
  })

  it('survives save and reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, '', PARAS)
    const again = await openPptx(await savePptx(opened))
    expect(getSlideNotesParagraphs(again.archive, again.deck.slides[0]!.path)).toEqual(PARAS)
  })

  it('plain notes read as unformatted paragraphs', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'a\nb')
    expect(getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)).toEqual([
      { runs: [{ text: 'a' }] },
      { runs: [{ text: 'b' }] },
    ])
  })
})
