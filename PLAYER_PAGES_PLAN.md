# Player Pages — Implementation Plan

**Status:** Drafted. Not yet implemented.
**Owner:** Benjamin Sleeter
**Date drafted:** 2026-04-24

This document captures the agreed scope for adding per-player profile pages
to the Program History section of the Gig Harbor Tides Baseball site. It is
self-contained — read this file before starting implementation in a future
session.

---

## Goal

Add `/history/player/[playerId]` — one statically-rendered profile page per
player in the program registry (~295 pages). Wire every player-name reference
across the site to link there.

---

## Current state of the data

- `gh-baseball-site/src/data/programHistory.json` already contains:
  - `players` registry with `playerId`, `displayName`, `years[]`, `grades{}`
  - per-season `battingLines` and `pitchingLines` (each tagged with `playerId`)
  - per-season `roster`, `record`, `headCoach`, championship flags
- Career roll-ups are already exposed by `getCareerBatting()` /
  `getCareerPitching()` (with `era` filter).
- Single-season + career leader functions exist:
  `singleSeasonBattingLeaders`, `singleSeasonPitchingLeaders`,
  `buildCareerHallOfFame`.
- Resolver merges callups across years (post-2026-04-24 enhancement) — players
  whose stat-line in year X precedes their roster appearance in year X±1
  are now linked correctly.

---

## Page anatomy

```
─────────────────────────────────────────────────────
PROGRAM HISTORY · PLAYER PROFILE      [● ACTIVE if 2026]
TOM FRIEDMAN
1992–95 · 4 seasons · Pre-BBCOR Era
[Career ribbon: 162 H · 6 HR · 82 RBI · .345 AVG]
─────────────────────────────────────────────────────

CAREER HIGHLIGHTS  (★ records held, tied, and top-N)
┌──────────────────┬──────────────────┬──────────────────┐
│ ★ PROGRAM RECORD │ #2 ALL-TIME      │ ★ TIED RECORD    │
│ Career Hits 162  │ Career RBIs 82   │ Single-Season    │
│                  │                  │ Triples · 5 ('95)│
└──────────────────┴──────────────────┴──────────────────┘
+ Compact "top 10 all-time in N more categories" list

CAREER BATTING
[Career totals row + per-season rows · sortable like the
 year-page Team Batting table]

CAREER PITCHING (only if any IP)
[Same idea]

SEASON-BY-SEASON
[1992 card] [1993 card] [1994 card] [1995 card]
   each clickable → /history/{year}
   shows: that year's team record, this player's headline
   stats, the player's grade for that year

GRADE PROGRESSION (small footer)
1992: Freshman → 1993: Sophomore → 1994: Junior → 1995: Senior
```

---

## Phase 1 — Data helpers

Add to `programHistory.ts`:

### `getPlayerCareerLines(playerId)`
Returns `{ battingByYear: Record<year, BattingLine>, pitchingByYear: Record<year, PitchingLine> }`.

### `getPlayerCareerSummary(playerId, era?)`
Returns the career batting + pitching aggregate for that single player.
Reuses the existing `getCareerBatting`/`getCareerPitching` filtered by
playerId, with the same rate-stat formulas (career AVG/OBP/SLG, ERA, WHIP,
K/7, K/BB).

### `getPlayerCareerHighlights(playerId)` ⭐ key feature
The "résumé generator" — scans every batting + pitching category (single-
season AND career) and returns a list of records this player holds or
ranks in. Output shape:

```ts
interface PlayerHighlight {
  kind: "single-season" | "career";
  category: string;            // "Most Hits", "Lowest ERA", etc.
  rank: number;                // 1, 2, 3 (or N if outside top-3)
  pool: number;
  era: "all" | "bbcor";        // computed against
  display: string;              // formatted value e.g. ".538"
  yearOrSpan: string;           // "2010" or "1992-95"
  isRecord: boolean;
  isTie: boolean;
  /** True for currently-active players whose career value is still growing. */
  inProgress: boolean;
}
```

Show order on the page:
1. Program records held (`isRecord && !inProgress`)
2. Tied program records
3. Active "in-progress" leadership (current player on pace)
4. Top-3 all-time finishes
5. Compact tail: "Top 10 in N more categories"

Computed for both `era: "all"` AND `era: "bbcor"` so the page can show
both contexts (or default to all-time, with a quiet "+N BBCOR-only" badge
when the BBCOR-era list adds more).

---

## Phase 2 — The page

`/history/player/[playerId]/page.tsx` (server component, SSG via
`generateStaticParams` listing every playerId in the registry).

Layout reuses existing components where possible:
- `PageHeader` (kicker, title, subtitle, stats ribbon)
- `EditorialDivider`
- `SectionHeader` for each section
- Card patterns from `HallOfFame.tsx` for the highlights section
- Table patterns from `StatTables.tsx` (sortable) for career batting/pitching
- New `PlayerSeasonCard` for season-by-season grid

### Header treatment
- Active badge: emerald "● ACTIVE" pill if `years.includes(LATEST_YEAR)`
- Era badge: small Pre-BBCOR / BBCOR pill matching the year-card style
- Stats ribbon picks 4 most relevant career stats (different for batters
  vs pitchers — pitchers default to W-L · IP · ERA · K)

### Highlights section
Maximum visual impact section. Mirrors the year-page HoF card aesthetic
(big display value, kicker label, gold trim for actual records, amber
"★ Tied Record" pill).

### Season-by-season cards
Mini-card per season with:
- Year (display font, large)
- Team record + championship pill if applicable
- Player's grade ("Senior", "Sophomore", etc.)
- Headline stats from that year (e.g. "32 H · .356 AVG" for a hitter,
  "22 IP · 0.00 ERA" for a pitcher; pick the most flattering 3)
- Hover/click → `/history/{year}`

---

## Phase 3 — Cross-linking

Wire every player-name reference site-wide to link to the new pages.

Touch-points (in priority order):

1. **`/history/[year]` Roster grid** — `RosterGrid` in
   `gh-baseball-site/src/app/history/[year]/page.tsx`
2. **`/history/[year]` Team Batting + Pitching tables** — sticky Player
   column in `StatTables.tsx`
3. **`/history/[year]` HoF cards** — holder names in `HallOfFame.tsx`,
   including the "Tied with" overflow list
4. **`/history/records` leaderboards** — both Single-Season and Career
   tabs in `app/history/records/page.tsx`
5. **`/history/career-leaders` cards** — leader + runner-up names in
   `app/history/career-leaders/CareerHallOfFame.tsx`
6. **(future) `/roster` current-team page** — if it lists player names,
   they should link too

Apply the link consistently using a tiny helper component:

```tsx
function PlayerLink({ playerId, fallback, children }) {
  if (!playerId) return <>{children ?? fallback}</>;
  return <Link href={`/history/player/${playerId}`} className="...">{children ?? fallback}</Link>;
}
```

Style: subtle hover-underline, no aggressive color change (don't break the
serif typography).

---

## Open decisions to confirm before starting

1. **Route**: Confirmed `/history/player/[playerId]`. (Alternatives
   considered: `/players/[id]` top-level, `/history/p/[id]` shorter — went
   with the namespaced version for clarity and to match the pattern of
   `/history/career-leaders`.)
2. **Coverage**: Generate pages for **all 295 players**, including 1-game
   call-ups. Sparse pages are accurate to the source data.
3. **Highlights depth**: Show top-3 finishes prominently; condense top-10
   into a one-line trailer ("Also top-10 all-time in: [Most BB, Most SB,
   Most TB]").
4. **Active flagging**: Use the existing `isActivePlayer(playerId)` helper.
   Add an "as of {LATEST_YEAR}" caption to highlight cards for active
   players so values aren't misread as final.
5. **Era toggle on highlights**: Default to All-Time. If BBCOR-era
   highlights add new entries, surface them as a smaller "+N BBCOR" badge
   on the section header (don't add another pill toggle — keep the page
   feeling like a profile, not another leaderboard view).
6. **Season-card era pill**: Re-use the same era pill (Pre-BBCOR /
   BBCOR) shown on the year-card landing page.

---

## Phase ordering when implementing

1. Phase 1 helpers (`getPlayerCareerLines`, `getPlayerCareerSummary`,
   `getPlayerCareerHighlights`) — purely data layer, no UI.
2. Phase 2 page implementation. Build the page using mock data on a single
   playerId first (e.g. Tom Friedman or Andy Cherbas — multi-season careers
   with both batting and pitching).
3. Phase 3 cross-linking. Add the `PlayerLink` helper, then apply it to
   each touch-point in the priority order above.
4. Verify production build pre-renders every player page and the page
   weight is reasonable.

---

## Risks / things to watch

- **Build time**: 295 pages × Phase 1 helpers running per page could be
  slow if helpers iterate every season. Mitigation: compute leaderboards
  once at module load and reuse, or memoize across renders.
- **Page-weight bloat**: Highlights section could get long for
  multi-record holders (Tom Friedman, Spencer Manjarrez). Cap at ~6
  prominent cards + condensed tail.
- **Sparse pages**: A 1-game callup will have an embarrassing-looking
  Highlights section ("not in top 10 of any category"). Acceptable —
  hide the section entirely when no highlights exist.
- **Name display consistency**: Confirm `displayNameOf` is used everywhere
  the canonical name is needed (not just raw `line.player`).

---

## Future extensions (not in scope for v1)

- Photo per player (require photo asset workflow)
- Position(s) played — would need source data we don't have
- Game-by-game splits — depends on having game logs (post-2023 only)
- Comparison view (compare two players side-by-side)
- Family/lineage callouts (Friedman brothers, Toglia trio)
- Coach pages (Pete Jansen / Shane Hannon / Ben Sleeter career arcs)
