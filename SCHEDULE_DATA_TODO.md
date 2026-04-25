# Schedule Data — Outstanding Work

**Status:** open work item
**Owner:** Benjamin Sleeter
**Date noted:** 2026-04-24

The site's per-year archive only has full schedule data (with opponents and
scores) for 8 of 34 years. For every other season we know the team's W-L
record and can derive single-game peaks from the Team Highlights sheet, but
we don't have game-by-game opponent + score data on file.

Game logs are valuable because they unlock:

- Per-game schedule rendering on each year page (re-enable the section we
  removed when coverage was too thin)
- Computed single-game records (Most Runs, Widest Margin, K's, etc.) so
  the Team Records page doesn't have to depend on each year's Team
  Highlights sheet — and so we can populate those records for years that
  never published one.
- Game-result detail on player pages once those exist.
- Streak verification (longest winning/losing streaks computed from data
  rather than from typed annual highlights).

---

## Coverage as of 2026-04-24

### Full schedule available (8 years)
- 2005, 2013, 2014, 2015, 2016, 2017, 2018, 2019

### Schedule missing (26 years)
- 1990, 1991 — no W-L record either; need yearbook research
- 1992 (16 games), 1993 (18), 1994 (20), 1995 (22), 1996 (22)
- 1997 (25 — state championship year), 1998 (21), 1999 (19)
- 2000 (16), 2001 (20), 2002 (18), 2003 (21), 2004 (20)
- 2006 (21), 2007 (21), 2008 (20), 2009 (20)
- 2010 (24), 2011 (24), 2012 (21)
- 2023 (23), 2024 (23), 2025 (28 — league champ + 3rd at state)
- 2026 (17 — current season, in progress)

(Numbers in parens are total games implied by the W-L record.)

---

## Where to source the data

- **2023-2026** (GameChanger era) — exportable from GameChanger directly.
  Each game has opponent + final score + box score. Should be
  straightforward to add.
- **2013-2019** — already done, sourced from typed schedule pages in the
  yearbook archive.
- **Pre-2013** — likely needs to be reconstructed from yearbooks, news
  archives (Gateway / Peninsula Gateway), or official KingCo / SPSL
  records. Higher effort.

---

## Implementation notes when data lands

When schedules are added (any era), the data path is:

1. Add a Schedule sheet to that year's `Historical/{year}/{year}_Season_Stats.xlsx`
   with columns: Date, Loc (Home/Away), Opponent, W/L, Score (`GH-Opp` form).
2. Re-run `python3 scripts/build_master_data.py`.
3. Re-enable the Schedule section on `/history/[year]` (currently commented
   out — was removed when coverage was too thin to be useful).
4. Build a "single-game records" computer that scans every year's schedule
   and computes Most Runs (Game), Widest Margin of Victory, etc. directly,
   replacing the highlights-sheet-based aggregation.

Until then, single-game records can be added manually via the
`highlights` override in `scripts/data/team_seasons.json` — same pattern
used for 2025 (Longest Winning Streak 16) and 2026 (Most Runs Game 27 vs
Mt. Tahoma).
