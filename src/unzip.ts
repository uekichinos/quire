import { strFromU8, Unzip, UnzipInflate } from 'fflate'
import { QuireError } from './errors'

export class XlsxReadError extends QuireError {
  constructor(message: string) {
    super(message)
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
  /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedstrings\.xml|styles\.xml|theme\/theme[0-9]*\.xml|worksheets\/[^/]+\.xml)$/i

function normalise(name: string): string {
  return name.replace(/\\/g, '/').replace(/^\.\//, '')
}

function concat(chunks: Uint8Array[], length: number): Uint8Array {
  const out = new Uint8Array(length)
  let offset = 0
  for (const c of chunks) {
    out.set(c, offset)
    offset += c.length
  }
  return out
}

/**
 * Extracts the allow-listed OOXML parts from an `.xlsx` (a ZIP).
 *
 * Uses fflate's **streaming** unzip so decompression can be aborted the moment a
 * part (or the workbook total) exceeds its cap — a ZIP directory that lies about
 * a part's size cannot force a large allocation. Non-allow-listed entries and
 * anything with `..` / an absolute path are never decompressed at all.
 */
export function extractParts(
  bytes: Uint8Array,
  limits: UnzipLimits = DEFAULT_LIMITS,
): Map<string, string> {
  const parts = new Map<string, string>()
  let total = 0
  let failure: XlsxReadError | null = null

  const unzip = new Unzip()
  unzip.register(UnzipInflate)

  unzip.onfile = (file) => {
    if (failure) return
    const name = normalise(file.name)
    if (name.includes('..') || name.startsWith('/') || !ALLOWED_PART.test(name)) return

    const chunks: Uint8Array[] = []
    let size = 0
    file.ondata = (err, chunk, final) => {
      if (failure) return
      if (err) {
        failure = new XlsxReadError(`corrupt part "${name}" (${err.message})`)
        return
      }
      size += chunk.length
      total += chunk.length
      if (size > limits.maxPartBytes) {
        failure = new XlsxReadError(`part "${name}" exceeds the ${limits.maxPartBytes}-byte limit`)
        return
      }
      if (total > limits.maxTotalBytes) {
        failure = new XlsxReadError(`workbook exceeds the ${limits.maxTotalBytes}-byte limit`)
        return
      }
      chunks.push(chunk)
      if (final) parts.set(name.toLowerCase(), strFromU8(concat(chunks, size)))
    }
    file.start()
  }

  try {
    unzip.push(bytes, true)
  } catch (err) {
    if (failure) throw failure
    if (err instanceof XlsxReadError) throw err
    throw new XlsxReadError(`not a readable .xlsx file (${(err as Error).message})`)
  }
  if (failure) throw failure

  if (!parts.has('xl/workbook.xml')) {
    throw new XlsxReadError('not a valid .xlsx — xl/workbook.xml is missing')
  }
  return parts
}
