import { attr, escapeText, XML_DECLARATION } from './xml'

/* -------------------------------------------------------------------------- */
/*  Namespaces / content types                                                 */
/* -------------------------------------------------------------------------- */

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
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

export function contentTypesXml(sheetCount: number, hasSharedStrings: boolean): string {
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
  return (
    `${XML_DECLARATION}\n<Types xmlns="${NS_CT}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
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

export function workbookXml(
  sheetNames: string[],
  filterRanges: readonly (string | undefined)[] = [],
): string {
  const sheets = sheetNames
    .map(
      (name, i) =>
        `<sheet${attr('name', name)} sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
    )
    .join('')

  const definedNames = sheetNames
    .map((name, i) =>
      filterRanges[i]
        ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">` +
          `${escapeText(absoluteSheetRange(name, filterRanges[i]!))}</definedName>`
        : '',
    )
    .join('')

  return (
    `${XML_DECLARATION}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
    `<sheets>${sheets}</sheets>` +
    (definedNames ? `<definedNames>${definedNames}</definedNames>` : '') +
    `<calcPr calcId="0" fullCalcOnLoad="1"/>` +
    `</workbook>`
  )
}

/* -------------------------------------------------------------------------- */
/*  Shared strings                                                             */
/* -------------------------------------------------------------------------- */

export function sharedStringsXml(strings: readonly string[], total: number): string {
  const items = strings
    .map((s) => `<si><t xml:space="preserve">${escapeText(s)}</t></si>`)
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
}

/**
 * Assembles a worksheet part. Child order follows the CT_Worksheet schema:
 * `dimension → sheetViews → cols → sheetData → autoFilter → mergeCells`.
 */
export function worksheetXml(s: WorksheetSections): string {
  return (
    `${XML_DECLARATION}\n<worksheet xmlns="${NS_MAIN}">` +
    `<dimension ref="${s.dimension}"/>` +
    (s.sheetViews ?? '') +
    (s.cols ?? '') +
    `<sheetData>${s.rows}</sheetData>` +
    (s.autoFilter ?? '') +
    (s.mergeCells ?? '') +
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
