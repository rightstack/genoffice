# Table ops

> Edit an existing table element: cell text, merges, row/column structure, row heights, column widths, cell anchors, style presets and borders.

All ops take `target:{slide, el}` where `el` is the table's id (type `table` in
the outline). `row` and `col` are 0-based. Chart edits (`setChart`) live in the
`edit_chart` tool.

### setTableCell

`{row,col,paragraphs}`

Replaces one cell's text. Runs you do not restyle inherit the cell's current
formatting, so plain `{text}` runs keep size, color and bold.

| Field      | Type                                                                                            | Notes                   |
| ---------- | ----------------------------------------------------------------------------------------------- | ----------------------- |
| row, col   | integer                                                                                         | 0-based                 |
| paragraphs | array of `{runs:[{text, bold?, italic?, fontSize?, color?}], align?, bullet?, lineSpacingPct?}` | Same shape as `setText` |

```json
{
  "op": "setTableCell",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 0,
  "paragraphs": [{ "runs": [{ "text": "Metric", "bold": true }], "align": "center" }]
}
```

Common mistakes

- Passing the cell text as a string: it must be a paragraph array.
- Row or column outside the grid: the op reports the failing (row, col).

### tableMerge

`{kind:"merge-right"|"merge-down"|"split",row,col}`

Merges the cell at (row, col) with its right or lower neighbor, or splits a
merged cell back into its grid cells.

```json
{
  "op": "tableMerge",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "merge-right",
  "row": 0,
  "col": 0
}
```

Common mistakes

- Merging across an existing merge boundary: the op refuses; split first.

### tableStructure

`{kind:"insert-row"|"delete-row"|"insert-col"|"delete-col",index,before?}`

Inserts or deletes a row or column. Insert goes after `index` unless
`before:true`.

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "insert-row",
  "index": 1
}
```

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "delete-col",
  "index": 0
}
```

Common mistakes

- Tables with merged cells refuse row/column surgery: split the merges first (`tableMerge` with `kind:"split"`).

### tableMove

`{axis:"row"|"col",from,to}`

Moves one row or column to another boundary in one step: `to` is the boundary
it lands on (`0` = before the first, the row/column count = after the last).
Text, cell formatting, widths and heights travel with it.

```json
{ "op": "tableMove", "target": { "slide": 0, "el": "e_TABLE" }, "axis": "row", "from": 0, "to": 2 }
```

Common mistakes

- `to` is a boundary, not the final index: moving row 0 below row 1 is `to: 2`.
- Tables with merged cells refuse the move, like `tableStructure`.

### setTableRowHeight

`{row,hEmu}`

Sets one row's height in EMU (the table frame grows or shrinks accordingly).

```json
{ "op": "setTableRowHeight", "target": { "slide": 0, "el": "e_TABLE" }, "row": 0, "hEmu": 457200 }
```

### setTableCellAnchor

`{row,col,anchor:"top"|"middle"|"bottom"}`

Vertical alignment of the text inside one cell.

```json
{
  "op": "setTableCellAnchor",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 1,
  "anchor": "middle"
}
```

### setTableColWidth

`{col,wEmu}`

Sets one column's width in EMU.

```json
{ "op": "setTableColWidth", "target": { "slide": 0, "el": "e_TABLE" }, "col": 0, "wEmu": 2743200 }
```

### setTableStyle

`{styleName} | {styleId?,firstRow?,lastRow?,firstCol?,lastCol?,bandRow?,bandCol?,keepFormatting?,shadingColor?,borderColor?,borderWidthPt?,borderPreset?,cells?,cellSides?,cellMargins?,cellVert?,clearTextFormatting?}`

Restyles the whole table: apply one fixed-color preset by `styleName`, pick one
of PowerPoint's 74 built-in styles by `styleId` (gallery name such as
`"Medium Style 2 - Accent 1"` or its GUID; they follow the deck's theme
colors), or change individual region flags, cell shading and border lines. A
preset wins over the other fields; a preset or `styleId` clears direct cell
fills and borders like PowerPoint's style gallery (`keepFormatting: true`
keeps them). `slides read` shows a table's current `style`.

| Field                               | Type                  | Notes                                                                                                                                                                                                                                          |
| ----------------------------------- | --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| styleName                           | string                | `none`, `lightGrid`, `zebraBlue`, `zebraGray`, `headerDarkBlue`, `headerOrange`, `noBorder`, `fullBorder`, `lilyGrid`                                                                                                                                    |
| styleId                             | string                | Built-in gallery name or `{GUID}`: `No Style, No Grid`, `No Style, Table Grid`, `Themed Style 1/2 - Accent N`, `Light Style 1/2/3 [- Accent N]`, `Medium Style 1/2/3/4 [- Accent N]`, `Dark Style 1 [- Accent N]`, `Dark Style 2 [- Accent N]` |
| firstRow                            | boolean               | Header-row emphasis                                                                                                                                                                                                                            |
| lastRow, firstCol, lastCol, bandCol | boolean               | Total row, first/last column emphasis, banded columns                                                                                                                                                                                          |
| bandRow                             | boolean               | Banded rows                                                                                                                                                                                                                                    |
| keepFormatting                      | boolean               | With `styleId`: keep direct cell fills/borders instead of clearing them                                                                                                                                                                        |
| shadingColor                        | `#RRGGBB` or `"none"` | Cell fill for every cell                                                                                                                                                                                                                       |
| borderColor                         | `#RRGGBB`             | Border line color                                                                                                                                                                                                                              |
| borderWidthPt                       | number (pt)           | Border line width; 0 < pt <= 1584                                                                                                                                                                                                              |
| borderPreset                        | `"all"` or `"none"`   | Draw all border lines, or clear them                                                                                                                                                                                                           |
| cells                               | `[{row,col}]`         | Limit shading/borderPreset to these cells (col = tc index)                                                                                                                                                                                     |
| cellSides                           | `[{row,col,side,ln}]` | One edge per entry: `side` `l`/`r`/`t`/`b`, `ln` `{color:"#RRGGBB",widthPt,dash?,compound?}` or `null` for no line                                                                                                                             |
| cellMargins                         | `{l?,r?,t?,b?}` (pt)  | Cell text insets, 0–144 pt; limited to `cells` when given                                                                                                                                                                                      |
| cellVert                            | string                | Cell text direction `horz`, `vert` (90°) or `vert270`; limited to `cells` when given                                                                                                                                                           |
| clearTextFormatting                 | boolean               | Drop direct character formatting (bold, italic, size, color, typeface) from every run in the table                                                                                                                                             |

```json
{ "op": "setTableStyle", "target": { "slide": 0, "el": "e_TABLE" }, "styleName": "zebraBlue" }
```

```json
{
  "op": "setTableStyle",
  "target": { "slide": 0, "el": "e_TABLE" },
  "styleId": "Medium Style 2 - Accent 1",
  "firstRow": true,
  "bandRow": true
}
```

```json
{
  "op": "setTableStyle",
  "target": { "slide": 0, "el": "e_TABLE" },
  "firstRow": true,
  "borderPreset": "all",
  "borderColor": "#BFBFBF",
  "borderWidthPt": 1
}
```

Common mistakes

- Passing a color name (`"blue"`): colors are `#RRGGBB`.
- Mixing `styleName` with the other fields: the preset is applied and the rest is ignored.
- A built-in name in `styleName`: gallery names and GUIDs go in `styleId`.
- Restyling one cell's text: that is `setTableCell` with styled runs, not this op.

### setChart (not-ai-callable)

`{patch:ChartEdit} — use the edit_chart tool instead`

Changes a chart's type, data, colors or elements. The `edit_chart` tool exposes
this with a validated schema.
