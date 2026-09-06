import { strFromU8, unzipSync, type UnzipFileInfo } from 'fflate'

export class XlsxReadError extends Error {
  constructor(message: string) {
    super(`@uekichinos/quire: ${message}`)
    this.name = 'XlsxReadError'
  }
}

export interface UnzipLimits {
  /** Max total uncompressed bytes across the parts we read. Default 100 MB. */
  maxTotalBytes: number
  /** Max uncompressed bytes for any single part. Default 50 MB. */
  maxPartBytes: number
}

export const DEFAULT_LIMITS: UnzipLimits = {
  maxTotalBytes: 100_000_000,
  maxPartBytes: 50_000_000,
}

/** Only these parts are ever extracted — anything else in the archive is ignored. */
const ALLOWED_PART =
  /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedstrings\.xml|styles\.xml|worksheets\/[^/]+\.xml)$/i

function normalise(name: string): string {
  return name.replace(/\\/g, '/').replace(/^\.\//, '')
}

/**
 * Extracts the allow-listed OOXML parts from an `.xlsx` (a ZIP), enforcing
 * size caps read from the central directory **before** decompression, then
 * re-checked after. Returns a map keyed by lower-cased part path.
 */
export function extractParts(
  bytes: Uint8Array,
  limits: UnzipLimits = DEFAULT_LIMITS,
): Map<string, string> {
  let declaredTotal = 0

  let raw: Record<string, Uint8Array>
  try {
    raw = unzipSync(bytes, {
      filter: (file: UnzipFileInfo): boolean => {
        const name = normalise(file.name)
        if (name.includes('..') || name.startsWith('/')) return false
        if (!ALLOWED_PART.test(name)) return false
        if (file.originalSize > limits.maxPartBytes) {
          throw new XlsxReadError(`part "${name}" exceeds the ${limits.maxPartBytes}-byte limit`)
        }
        declaredTotal += file.originalSize
        if (declaredTotal > limits.maxTotalBytes) {
          throw new XlsxReadError(`workbook exceeds the ${limits.maxTotalBytes}-byte limit`)
        }
        return true
      },
    })
  } catch (err) {
    if (err instanceof XlsxReadError) throw err
    throw new XlsxReadError(`not a readable .xlsx file (${(err as Error).message})`)
  }

  const parts = new Map<string, string>()
  let actualTotal = 0
  for (const [name, data] of Object.entries(raw)) {
    actualTotal += data.length
    if (actualTotal > limits.maxTotalBytes) {
      throw new XlsxReadError('workbook exceeds its size limit after decompression')
    }
    parts.set(normalise(name).toLowerCase(), strFromU8(data))
  }

  if (!parts.has('xl/workbook.xml')) {
    throw new XlsxReadError('not a valid .xlsx — xl/workbook.xml is missing')
  }
  return parts
}
