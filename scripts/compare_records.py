#!/usr/bin/env python3
"""Compare GameChanger single-season stats (2023-2025) against the
1990-2019 Hall of Fame records. Report whether any records have been broken
or approached."""
import csv
from pathlib import Path

STATS_DIR = Path(__file__).resolve().parent.parent / "Stats"
FILES = {
    2023: "Gig Harbor Varsity Tides Spring 2023 Stats.csv",
    2024: "Gig Harbor Tides Varsity Spring 2024 Stats.csv",
    2025: "Gig Harbor Varsity Tides Spring 2025 Stats.csv",
}

# Hall of Fame single-season records (1990-2019)
HOF = {
    "AVG":            {"val": 0.538,  "who": "Spencer Manjarrez", "note": "35/65 AB"},
    "OBP":            {"val": 0.677,  "who": "Spencer Manjarrez"},
    "H":              {"val": 47,     "who": "Tim Friedman"},
    "1B":             {"val": 28,     "who": "Cassidy Emery"},
    "2B":             {"val": 13,     "who": "Tim Friedman"},
    "3B":             {"val": 7,      "who": "Jordan Haworth"},
    "HR":             {"val": 7,      "who": "Friedman / Bigelow / Manjarrez"},
    "TB":             {"val": 83,     "who": "Tim Friedman"},
    "RBI":            {"val": 41,     "who": "Tim Friedman"},
    "R":              {"val": 45,     "who": "Tim Friedman"},
    "AB":             {"val": 114,    "who": "RJ Green"},
    "BB":             {"val": 32,     "who": "Aaron Araujo"},
    "HBP":            {"val": 14,     "who": "Cage Hardy"},
    "SB":             {"val": 26,     "who": "Spencer Manjarrez"},
    "W":              {"val": 10,     "who": "Matt Gardner"},
    "SV":             {"val": 5,      "who": "Anthony Gilich"},
    "IP":             {"val": 78.0,   "who": "Owen Wild"},
    "K_pit":          {"val": 112,    "who": "Owen Wild"},
    "ERA":            {"val": 0.62,   "who": "Michael Toglia", "note": "Min 5 decisions"},
}


def load_season(path):
    with path.open(newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    section_row = rows[0]
    header_row = rows[1]
    current = ""
    sections = []
    for s in section_row:
        if s.strip():
            current = s.strip()
        sections.append(current)
    keys = []
    for sec, name in zip(sections, header_row):
        name = name.strip()
        if sec and name not in ("Number", "Last", "First"):
            keys.append(f"{sec}:{name}")
        else:
            keys.append(name)
    players = []
    for row in rows[2:]:
        if not row or not row[0]:
            continue
        if row[0] == "Totals":
            continue
        if row[0] == "Glossary":
            break
        players.append(dict(zip(keys, row)))
    return players


def to_num(v, default=0.0):
    if v is None:
        return default
    v = v.strip()
    if v in ("", "-", "N/A", "Inf"):
        return default
    try:
        return float(v)
    except ValueError:
        return default


def ip_frac_to_thirds(ip_raw):
    whole, _, frac = ip_raw.partition(".")
    try:
        t = int(whole or 0) * 3
        if frac:
            t += int(frac[0])
        return t
    except ValueError:
        return 0


def full_name(rec):
    return f"{rec.get('First','').strip()} {rec.get('Last','').strip()}".strip()


def all_seasons():
    out = []
    for year, fname in FILES.items():
        for p in load_season(STATS_DIR / fname):
            out.append((year, p))
    return out


def best_across(stat_extractor, qualifier=None, reverse=True, top=3):
    """stat_extractor(year, player_rec) -> (value, summary_str) or None"""
    rows = []
    for year, p in all_seasons():
        result = stat_extractor(year, p)
        if result is None:
            continue
        if qualifier and not qualifier(year, p):
            continue
        val, summary = result
        rows.append((val, year, full_name(p), summary))
    rows.sort(reverse=reverse)
    return rows[:top]


def fmt_hof(key):
    h = HOF[key]
    v = h["val"]
    who = h["who"]
    note = h.get("note", "")
    if isinstance(v, float) and v < 1:
        vstr = f"{v:.3f}".lstrip("0")
    else:
        vstr = f"{v:.2f}" if isinstance(v, float) else str(v)
    return f"{vstr} — {who}" + (f" ({note})" if note else "")


def compare(label, hof_key, leaders, beats_fn, close_fn=None):
    hof_val = HOF[hof_key]["val"]
    print(f"\n{label}")
    print(f"  HoF Record: {fmt_hof(hof_key)}")
    print("  GameChanger-era best (top 3):")
    status = "no_threat"
    for v, yr, name, summary in leaders:
        marker = ""
        if beats_fn(v, hof_val):
            marker = "  *** BROKEN ***"
            status = "broken"
        elif close_fn and close_fn(v, hof_val) and status != "broken":
            marker = "  (within striking distance)"
            if status == "no_threat":
                status = "close"
        print(f"    {name} ({yr}):  {summary}{marker}")


def main():
    # ─── Batting ───
    compare(
        "SINGLE-SEASON AVG (min 20 PA)",
        "AVG",
        best_across(
            lambda y, p: (
                (to_num(p.get("Batting:H")) / to_num(p.get("Batting:AB")))
                if to_num(p.get("Batting:AB")) >= 1
                else None,
                f"{to_num(p.get('Batting:H'))/to_num(p.get('Batting:AB')):.3f} ({int(to_num(p.get('Batting:H')))}/{int(to_num(p.get('Batting:AB')))})"
            ) if to_num(p.get("Batting:AB")) >= 20 else None,
        ),
        beats_fn=lambda v, h: v > h,
        close_fn=lambda v, h: v >= h - 0.05,
    )

    compare(
        "SINGLE-SEASON OBP (min 20 PA)",
        "OBP",
        best_across(
            lambda y, p: (
                (
                    (to_num(p.get("Batting:H")) + to_num(p.get("Batting:BB")) + to_num(p.get("Batting:HBP")))
                    / max(1, to_num(p.get("Batting:AB")) + to_num(p.get("Batting:BB")) + to_num(p.get("Batting:HBP")) + to_num(p.get("Batting:SF")))
                ),
                f"OBP .{int(1000*(to_num(p.get('Batting:H'))+to_num(p.get('Batting:BB'))+to_num(p.get('Batting:HBP')))/max(1,to_num(p.get('Batting:AB'))+to_num(p.get('Batting:BB'))+to_num(p.get('Batting:HBP'))+to_num(p.get('Batting:SF')))):03d}"
            ) if to_num(p.get("Batting:AB")) >= 20 else None,
        ),
        beats_fn=lambda v, h: v > h,
        close_fn=lambda v, h: v >= h - 0.05,
    )

    # Simple counting stats
    for label, stat_key, hof_key in [
        ("SINGLE-SEASON HITS",          "Batting:H",   "H"),
        ("SINGLE-SEASON SINGLES",       "Batting:1B",  "1B"),
        ("SINGLE-SEASON DOUBLES",       "Batting:2B",  "2B"),
        ("SINGLE-SEASON TRIPLES",       "Batting:3B",  "3B"),
        ("SINGLE-SEASON HOME RUNS",     "Batting:HR",  "HR"),
        ("SINGLE-SEASON TOTAL BASES",   "Batting:TB",  "TB"),
        ("SINGLE-SEASON RBI",           "Batting:RBI", "RBI"),
        ("SINGLE-SEASON RUNS SCORED",   "Batting:R",   "R"),
        ("SINGLE-SEASON AT BATS",       "Batting:AB",  "AB"),
        ("SINGLE-SEASON WALKS",         "Batting:BB",  "BB"),
        ("SINGLE-SEASON HBP",           "Batting:HBP", "HBP"),
        ("SINGLE-SEASON STOLEN BASES",  "Batting:SB",  "SB"),
    ]:
        compare(
            label, hof_key,
            best_across(
                lambda y, p, k=stat_key: (int(to_num(p.get(k))), f"{int(to_num(p.get(k)))}"),
            ),
            beats_fn=lambda v, h: v > h,
            close_fn=lambda v, h: v >= h - max(2, 0.2*h),
        )

    # Pitching
    for label, stat_key, hof_key in [
        ("SINGLE-SEASON WINS",          "Pitching:W",  "W"),
        ("SINGLE-SEASON SAVES",         "Pitching:SV", "SV"),
        ("SINGLE-SEASON STRIKEOUTS (pitching)", "Pitching:SO", "K_pit"),
    ]:
        compare(
            label, hof_key,
            best_across(
                lambda y, p, k=stat_key: (int(to_num(p.get(k))), f"{int(to_num(p.get(k)))} · IP {p.get('Pitching:IP','0')}"),
            ),
            beats_fn=lambda v, h: v > h,
            close_fn=lambda v, h: v >= h - max(2, 0.2*h),
        )

    # IP (pitching)
    compare(
        "SINGLE-SEASON INNINGS PITCHED",
        "IP",
        best_across(
            lambda y, p: (
                ip_frac_to_thirds(p.get("Pitching:IP", "0")) / 3,
                f"{p.get('Pitching:IP','0')} IP"
            ),
        ),
        beats_fn=lambda v, h: v > h,
        close_fn=lambda v, h: v >= h - max(2, 0.15*h),
    )

    # ERA (min 5 decisions)
    def era_leader(y, p):
        t = ip_frac_to_thirds(p.get("Pitching:IP", "0"))
        if t == 0:
            return None
        w = to_num(p.get("Pitching:W"))
        l = to_num(p.get("Pitching:L"))
        decisions = w + l
        if decisions < 5:
            return None
        er = to_num(p.get("Pitching:ER"))
        ip = t / 3
        era = (er * 7) / ip
        return (era, f"ERA {era:.2f} · {int(w)}-{int(l)} · {p.get('Pitching:IP','0')} IP")

    compare(
        "SINGLE-SEASON ERA (min 5 decisions)",
        "ERA",
        best_across(era_leader, reverse=False),  # lowest is best
        beats_fn=lambda v, h: v < h,
        close_fn=lambda v, h: v <= h + 0.3,
    )

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Legend: *** BROKEN *** = GameChanger-era player exceeded the HoF record.")


if __name__ == "__main__":
    main()
