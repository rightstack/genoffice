/**
 * [lily] Speaker-notes body in the presenter view. Plain text here; a host that
 * stores notes as markup swaps this module to render it (one seam, no other change).
 */
import React from 'react'

export function NotesText({ text }: { text: string }): React.ReactElement {
  return <>{text}</>
}
