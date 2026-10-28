// The package's time zone (CR-018; the portfolio's date rule of 2026-10-28, docx4j CR-067 D12): an instant a document
// holds (a revision's or a comment's stamp, a date field) is shown, and a `w:date` without `dateUtc` read, in the zone
// the caller sets on the package it is working on, else the runtime's. Word writes such a `w:date` as the author's local
// wall-clock time labelled `Z` (CR-002 section 29, check 23), so the zone is what turns it into an instant and back.
// Under Office JS the runtime's zone is the client's; under Node on a server it is the server's, hence the option.
// Nothing here reaches a saved file as such: the zone decides which wall-clock fields are written for an instant.

/** The fields a wall clock shows: `month` 1 to 12, `hour` 0 to 23. */
export interface WallClock {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
}

/** The runtime's zone, an IANA id (`Australia/Sydney`; `UTC` where the runtime has no better answer). */
export function runtimeTimeZone(): string {
  return new Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
}

/**
 * `id` when the runtime's `Intl` knows it as a zone, else a `RangeError` naming it: a package's
 * `dateTimeZone` is refused at the set, never ignored, since a wrong zone shifts every stamp.
 */
export function assertTimeZone(id: string): string {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: id });
  } catch {
    throw new RangeError(`Not a time zone this runtime knows: ${JSON.stringify(id)} (an IANA id, such as "Australia/Sydney")`);
  }
  return id;
}

/** The wall clock `date` shows in `zone`; the runtime's zone where `zone` is undefined. */
export function wallClockIn(date: Date, zone: string | undefined): WallClock {
  if (zone === undefined) {
    return {
      year: date.getFullYear(), month: date.getMonth() + 1, day: date.getDate(),
      hour: date.getHours(), minute: date.getMinutes(), second: date.getSeconds(),
    };
  }
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: zone, hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(date);
  const field = (type: Intl.DateTimeFormatPartTypes): number => Number(parts.find((p) => p.type === type)?.value ?? NaN);
  return {
    year: field('year'), month: field('month'), day: field('day'),
    hour: field('hour') % 24, minute: field('minute'), second: field('second'), // hour12: false gives 24 at midnight on some engines
  };
}

/**
 * The instant at which a wall clock in `zone` shows `clock`; the runtime's zone where `zone` is
 * undefined. Found by offset: the clock taken as UTC, corrected by what the zone shows for that
 * guess, twice, which settles across a daylight-saving step; a clock time that does not exist on a
 * spring-forward day gives the instant an hour on, as `Date` does for the runtime's zone.
 */
export function instantOf(clock: WallClock, zone: string | undefined): Date {
  if (zone === undefined) return new Date(clock.year, clock.month - 1, clock.day, clock.hour, clock.minute, clock.second);
  const asUtc = Date.UTC(clock.year, clock.month - 1, clock.day, clock.hour, clock.minute, clock.second);
  let guess = asUtc;
  for (let i = 0; i < 2; i++) {
    const shown = wallClockIn(new Date(guess), zone);
    guess -= Date.UTC(shown.year, shown.month - 1, shown.day, shown.hour, shown.minute, shown.second) - asUtc;
  }
  return new Date(guess);
}
