/**
 * Excel 1900 date system. Serial `1` is 1900-01-01; the fractional part is the
 * time of day. Excel incorrectly treats 1900 as a leap year, so every serial
 * from 1900-03-01 onward is shifted up by one — we reproduce that so the numbers
 * match what Excel shows.
 *
 * Adapted from ExcelJS date handling (MIT).
 */

const EPOCH_1900_UTC = Date.UTC(1899, 11, 31)
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1)
const DAY_MS = 86_400_000

/** `Date → Excel serial number`. Uses the date's **local** calendar fields. */
export function dateToSerial(date: Date): number {
  const time = date.getTime()
  if (Number.isNaN(time)) {
    throw new Error('@uekichinos/quire: cannot write an invalid Date')
  }

  const dayStartUtc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  let days = Math.round((dayStartUtc - EPOCH_1900_UTC) / DAY_MS)

  if (days < 1) {
    throw new Error(
      '@uekichinos/quire: dates before 1900-01-01 are not representable in the 1900 date system',
    )
  }

  // Excel's phantom 1900-02-29.
  if (days >= 60) days += 1

  const msIntoDay =
    date.getHours() * 3_600_000 +
    date.getMinutes() * 60_000 +
    date.getSeconds() * 1_000 +
    date.getMilliseconds()

  return days + msIntoDay / DAY_MS
}

/**
 * `Excel serial number → Date` — the inverse of `dateToSerial`. The returned
 * `Date` carries the spreadsheet's calendar date/time in its **local** fields
 * (spreadsheets have no timezone), so it round-trips with `dateToSerial`.
 */
export function serialToDate(serial: number, date1904 = false): Date {
  if (!Number.isFinite(serial)) {
    throw new Error('@uekichinos/quire: invalid date serial')
  }

  let whole = Math.floor(serial)
  const frac = serial - whole

  if (!date1904 && whole >= 60) whole -= 1 // undo Excel's phantom 1900-02-29

  const dayUtc = (date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC) + whole * DAY_MS
  const d = new Date(dayUtc)
  const localMidnight = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())
  return new Date(localMidnight.getTime() + Math.round(frac * DAY_MS))
}
