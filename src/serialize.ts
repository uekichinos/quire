import { toArgb } from './style'
import type { RichText, RichTextRun } from './workbook'
import { attr, escapeText, XML_DECLARATION } from './xml'

/* -------------------------------------------------------------------------- */
/*  Namespaces / content types                                                 */
/* -------------------------------------------------------------------------- */

export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

const CT_WORKBOOK =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const CT_WORKSHEET =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_STYLES =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const CT_SST =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml'
const CT_APP = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'

export function contentTypesXml(
  sheetCount: number,
  hasSharedStrings: boolean,
  /** 1-based sheet numbers that have a `commentsN.xml` + `vmlDrawingN.vml` pair. */
  commentSheets: readonly number[] = [],
): string {
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="${CT_WORKBOOK}"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="${CT_STYLES}"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="${CT_CORE}"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="${CT_APP}"/>`,
  ]
  for (let i = 1; i <= sheetCount; i++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="${CT_WORKSHEET}"/>`,
    )
  }
  if (hasSharedStrings) {
    overrides.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_SST}"/>`)
  }
  for (const n of commentSheets) {
    overrides.push(`<Override PartName="/xl/comments${n}.xml" ContentType="${CT_COMMENTS}"/>`)
  }
  return (
    `${XML_DECLARATION}\n<Types xmlns="${NS_CT}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    (commentSheets.length
      ? `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`
      : '') +
    overrides.join('') +
    `</Types>`
  )
}

/* -------------------------------------------------------------------------- */
/*  Relationships                                                              */
/* -------------------------------------------------------------------------- */

const REL_CORE_PROPS =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'

export function rootRelsXml(): string {
  return (
    `${XML_DECLARATION}\n<Relationships xmlns="${NS_PKG_REL}">` +
    `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="${REL_CORE_PROPS}" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`
  )
}

export const REL_HYPERLINK = `${NS_R}/hyperlink`
export const REL_COMMENTS = `${NS_R}/comments`
export const REL_VML = `${NS_R}/vmlDrawing`

export interface WorksheetRelEntry {
  id: string
  /** Full relationship-type URI, e.g. `REL_HYPERLINK`. */
  type: string
  target: string
  /** `TargetMode="External"` — set for hyperlinks, omitted for in-package parts. */
  external?: boolean
}

/** A worksheet's own `_rels/sheetN.xml.rels` — hyperlinks and/or its comments + VML drawing. */
export function worksheetRelsXml(rels: readonly WorksheetRelEntry[]): string {
  const items = rels
    .map(
      (r) =>
        `<Relationship${attr('Id', r.id)} Type="${r.type}"${attr('Target', r.target)}` +
        `${r.external ? ' TargetMode="External"' : ''}/>`,
    )
    .join('')
  return `${XML_DECLARATION}\n<Relationships xmlns="${NS_PKG_REL}">${items}</Relationships>`
}

/** A worksheet's `xl/commentsN.xml` — classic (legacy) cell comments. */
export function commentsXml(
  authors: readonly string[],
  comments: readonly { ref: string; authorId: number; text: string }[],
): string {
  const authorsXml = authors.map((a) => `<author>${escapeText(a)}</author>`).join('')
  const list = comments
    .map(
      (c) =>
        `<comment${attr('ref', c.ref)} authorId="${c.authorId}">` +
        `<text><r><t xml:space="preserve">${escapeText(c.text)}</t></r></text></comment>`,
    )
    .join('')
  return (
    `${XML_DECLARATION}\n<comments xmlns="${NS_MAIN}">` +
    `<authors>${authorsXml}</authors><commentList>${list}</commentList></comments>`
  )
}

/**
 * The legacy VML drawing that anchors each comment's hover box — required alongside
 * `commentsN.xml` for Excel to render the indicator triangle and popup correctly.
 * `row`/`col` are 0-based.
 */
export function vmlDrawingXml(shapes: readonly { id: number; row: number; col: number }[]): string {
  const shapeXml = shapes
    .map(({ id, row, col }) => {
      const anchor = `${col + 1}, 15, ${row}, 2, ${col + 3}, 15, ${row + 4}, 4`
      return (
        `<v:shape id="_x0000_s${1000 + id}" type="#_x0000_t202" ` +
        `style='position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;` +
        `z-index:${id + 1};visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">` +
        `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>` +
        `<v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>` +
        `<x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
        `<x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`
      )
    })
    .join('')
  return (
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">` +
    `<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shapeXml +
    `</xml>`
  )
}

export function workbookRelsXml(sheetCount: number, hasSharedStrings: boolean): string {
  const rels: string[] = []
  let id = 1
  for (let i = 1; i <= sheetCount; i++) {
    rels.push(
      `<Relationship Id="rId${id++}" Type="${NS_R}/worksheet" Target="worksheets/sheet${i}.xml"/>`,
    )
  }
  rels.push(`<Relationship Id="rId${id++}" Type="${NS_R}/styles" Target="styles.xml"/>`)
  if (hasSharedStrings) {
    rels.push(
      `<Relationship Id="rId${id++}" Type="${NS_R}/sharedStrings" Target="sharedStrings.xml"/>`,
    )
  }
  return `${XML_DECLARATION}\n<Relationships xmlns="${NS_PKG_REL}">${rels.join('')}</Relationships>`
}

/* -------------------------------------------------------------------------- */
/*  Workbook                                                                   */
/* -------------------------------------------------------------------------- */

/** `"Sales"`, `"A1:E1"` → `'Sales'!$A$1:$E$1` (sheet-qualified, absolute). */
function absoluteSheetRange(sheetName: string, range: string): string {
  const abs = range.replace(/([A-Z]+)([0-9]+)/g, '$$$1$$$2')
  return `'${sheetName.replace(/'/g, "''")}'!${abs}`
}

/** A user-defined named range: `name` refers to `range` on `sheetName`. */
export interface DefinedNameEntry {
  name: string
  sheetName: string
  range: string
}

export function workbookXml(
  sheetNames: string[],
  filterRanges: readonly (string | undefined)[] = [],
  definedNames: readonly DefinedNameEntry[] = [],
  printAreas: readonly (string | undefined)[] = [],
): string {
  const sheets = sheetNames
    .map(
      (name, i) =>
        `<sheet${attr('name', name)} sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
    )
    .join('')

  const autoFilterNames = sheetNames
    .map((name, i) =>
      filterRanges[i]
        ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">` +
          `${escapeText(absoluteSheetRange(name, filterRanges[i]!))}</definedName>`
        : '',
    )
    .join('')

  const printAreaNames = sheetNames
    .map((name, i) =>
      printAreas[i]
        ? `<definedName name="_xlnm.Print_Area" localSheetId="${i}">` +
          `${escapeText(absoluteSheetRange(name, printAreas[i]!))}</definedName>`
        : '',
    )
    .join('')

  const userNames = definedNames
    .map(
      (d) =>
        `<definedName${attr('name', d.name)}>` +
        `${escapeText(absoluteSheetRange(d.sheetName, d.range))}</definedName>`,
    )
    .join('')

  const allNames = autoFilterNames + printAreaNames + userNames

  return (
    `${XML_DECLARATION}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
    `<sheets>${sheets}</sheets>` +
    (allNames ? `<definedNames>${allNames}</definedNames>` : '') +
    `<calcPr calcId="0" fullCalcOnLoad="1"/>` +
    `</workbook>`
  )
}

/* -------------------------------------------------------------------------- */
/*  Shared strings                                                             */
/* -------------------------------------------------------------------------- */

function runXml(run: RichTextRun): string {
  const f = run.font
  const rPr = f
    ? '<rPr>' +
      (f.bold ? '<b/>' : '') +
      (f.italic ? '<i/>' : '') +
      (f.underline ? '<u/>' : '') +
      (f.size != null ? `<sz val="${f.size}"/>` : '') +
      (f.color ? `<color rgb="${toArgb(f.color)}"/>` : '') +
      (f.name ? `<rFont${attr('val', f.name)}/>` : '') +
      '</rPr>'
    : ''
  return `<r>${rPr}<t xml:space="preserve">${escapeText(run.text)}</t></r>`
}

export function sharedStringsXml(strings: readonly (string | RichText)[], total: number): string {
  const items = strings
    .map((s) =>
      typeof s === 'string'
        ? `<si><t xml:space="preserve">${escapeText(s)}</t></si>`
        : `<si>${s.map(runXml).join('')}</si>`,
    )
    .join('')
  return (
    `${XML_DECLARATION}\n<sst xmlns="${NS_MAIN}" count="${total}" uniqueCount="${strings.length}">` +
    items +
    `</sst>`
  )
}

/* -------------------------------------------------------------------------- */
/*  Worksheet                                                                  */
/* -------------------------------------------------------------------------- */

export interface WorksheetSections {
  /** `<sheetPr>…</sheetPr>` (fit-to-page flag) — must precede `dimension`, or omitted. */
  sheetPr?: string
  dimension: string
  /** Full `<sheetViews>…</sheetViews>` (freeze panes), or omitted. */
  sheetViews?: string
  /** Full `<cols>…</cols>`, or omitted. */
  cols?: string
  /** The `<row>…` body that goes inside `<sheetData>`. */
  rows: string
  /** `<autoFilter ref="…"/>`, or omitted. */
  autoFilter?: string
  /** Full `<mergeCells>…</mergeCells>`, or omitted. */
  mergeCells?: string
  /** One or more full `<conditionalFormatting>…</conditionalFormatting>` blocks, or omitted. */
  conditionalFormatting?: string
  /** Full `<dataValidations>…</dataValidations>`, or omitted. */
  dataValidations?: string
  /** Full `<hyperlinks>…</hyperlinks>`, or omitted. */
  hyperlinks?: string
  /** `<pageMargins .../>`, or omitted. */
  pageMargins?: string
  /** `<pageSetup .../>`, or omitted. */
  pageSetup?: string
  /** `<legacyDrawing r:id="…"/>` (comments' VML anchor), or omitted. */
  legacyDrawing?: string
}

/**
 * Assembles a worksheet part. Child order follows the CT_Worksheet schema:
 * `sheetPr → dimension → sheetViews → cols → sheetData → autoFilter → mergeCells →
 * conditionalFormatting → dataValidations → hyperlinks → pageMargins → pageSetup →
 * legacyDrawing`.
 */
export function worksheetXml(s: WorksheetSections): string {
  return (
    `${XML_DECLARATION}\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
    (s.sheetPr ?? '') +
    `<dimension ref="${s.dimension}"/>` +
    (s.sheetViews ?? '') +
    (s.cols ?? '') +
    `<sheetData>${s.rows}</sheetData>` +
    (s.autoFilter ?? '') +
    (s.mergeCells ?? '') +
    (s.conditionalFormatting ?? '') +
    (s.dataValidations ?? '') +
    (s.hyperlinks ?? '') +
    (s.pageMargins ?? '') +
    (s.pageSetup ?? '') +
    (s.legacyDrawing ?? '') +
    `</worksheet>`
  )
}

/* -------------------------------------------------------------------------- */
/*  Doc props                                                                  */
/* -------------------------------------------------------------------------- */

export function coreXml(iso: string): string {
  return (
    `${XML_DECLARATION}\n<cp:coreProperties ` +
    `xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
    `xmlns:dcterms="http://purl.org/dc/terms/" ` +
    `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:creator>@uekichinos/quire</dc:creator>` +
    `<cp:lastModifiedBy>@uekichinos/quire</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>` +
    `</cp:coreProperties>`
  )
}

export function appXml(): string {
  return (
    `${XML_DECLARATION}\n<Properties ` +
    `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
    `<Application>@uekichinos/quire</Application>` +
    `</Properties>`
  )
}
