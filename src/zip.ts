import { strToU8, zipSync, type Zippable } from 'fflate'

/** Fixed timestamp so identical input produces identical output (golden-file friendly). */
const FIXED_MTIME = new Date('2020-01-01T00:00:00Z')

/**
 * Packs the given part path → content map into a ZIP container (the `.xlsx`
 * envelope). Strings are encoded as UTF-8. `[Content_Types].xml` is stored
 * uncompressed as some readers expect; everything else is DEFLATEd.
 */
export function zipParts(parts: Record<string, string | Uint8Array>): Uint8Array {
  const zippable: Zippable = {}
  for (const [path, content] of Object.entries(parts)) {
    const bytes = typeof content === 'string' ? strToU8(content) : content
    const stored = path === '[Content_Types].xml'
    zippable[path] = [bytes, { level: stored ? 0 : 6, mtime: FIXED_MTIME }]
  }
  return zipSync(zippable)
}
