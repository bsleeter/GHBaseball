#!/usr/bin/env python3
"""Summarize Gig Harbor Varsity Tides stats across 2023, 2024, and 2025 seasons."""
import csv
from pathlib import Path

STATS_DIR = Path(__file__).resolve().parent.parent / "Stats"

FILES = {
    2023: "Gig Harbor Varsity Tides Spring 2023 Stats.csv",
    2024: "Gig Harbor Tides Varsity Spring 2024 Stats.csv",
    2025: "Gig Harbor Varsity Tides Spring 2025 Stats.csv",
}

# Header row 2 has the actual column names. Row 1 is section labels (Batting/Pitching/Fielding).
# We need to disambiguate columns that share names across sections (e.g. GP, H, R, BB, SO, CS, SB, SB%, PIK, LOB, LD%, FB%, GB%, BABIP, BA/RISP).


def load_season(path: Path):
    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    section_row = rows[0]
    header_row = rows[1]
    # Carry-forward section label
    sections = []
    current = ""
    for s in section_row:
        if s.strip():
            current = s.strip()
        sections.append(current)
    # Build unique column keys like "Batting:H", "Pitching:H"
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
            totals = dict(zip(keys, row))
            continue
        if row[0] == "Glossary":
            break
        rec = dict(zip(keys, row))
        # Skip any empty-number rows
        players.append(rec)
    return players, totals, keys


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


def fmt_avg(x):
    # Baseball style: drop leading zero
    s = f"{x:.3f}"
    if s.startswith("0"):
        s = s[1:]
    elif s.startswith("-0"):
        s = "-" + s[2:]
    return s


def bat_line(rec):
    return {
        "GP": to_num(rec.get("Batting:GP")),
        "PA": to_num(rec.get("Batting:PA")),
        "AB": to_num(rec.get("Batting:AB")),
        "H":  to_num(rec.get("Batting:H")),
        "1B": to_num(rec.get("Batting:1B")),
        "2B": to_num(rec.get("Batting:2B")),
        "3B": to_num(rec.get("Batting:3B")),
        "HR": to_num(rec.get("Batting:HR")),
        "RBI": to_num(rec.get("Batting:RBI")),
        "R":  to_num(rec.get("Batting:R")),
        "BB": to_num(rec.get("Batting:BB")),
        "SO": to_num(rec.get("Batting:SO")),
        "HBP": to_num(rec.get("Batting:HBP")),
        "SF":  to_num(rec.get("Batting:SF")),
        "SB":  to_num(rec.get("Batting:SB")),
        "TB":  to_num(rec.get("Batting:TB")),
    }


def pit_line(rec):
    # GameChanger uses baseball IP format: 5.1 means 5 and 1/3 innings.
    ip_raw = rec.get("Pitching:IP", "0")
    ip_whole, _, frac = ip_raw.partition(".")
    try:
        ip_thirds = int(ip_whole or 0) * 3
        if frac:
            ip_thirds += int(frac[0])
    except ValueError:
        ip_thirds = 0
    return {
        "IP_thirds": ip_thirds,
        "BF": to_num(rec.get("Pitching:BF")),
        "H":  to_num(rec.get("Pitching:H")),
        "R":  to_num(rec.get("Pitching:R")),
        "ER": to_num(rec.get("Pitching:ER")),
        "BB": to_num(rec.get("Pitching:BB")),
        "SO": to_num(rec.get("Pitching:SO")),
        "W":  to_num(rec.get("Pitching:W")),
        "L":  to_num(rec.get("Pitching:L")),
        "SV": to_num(rec.get("Pitching:SV")),
        "HBP": to_num(rec.get("Pitching:HBP")),
        "#P": to_num(rec.get("Pitching:#P")),
    }


def ip_from_thirds(t):
    whole = t // 3
    rem = t % 3
    return f"{int(whole)}.{int(rem)}"


def calc_batting_rates(b):
    ab, h, bb, hbp, sf, tb = b["AB"], b["H"], b["BB"], b["HBP"], b["SF"], b["TB"]
    avg = h / ab if ab else 0
    obp_den = ab + bb + hbp + sf
    obp = (h + bb + hbp) / obp_den if obp_den else 0
    slg = tb / ab if ab else 0
    return avg, obp, slg, obp + slg


def calc_pitching_rates(p):
    ip = p["IP_thirds"] / 3 if p["IP_thirds"] else 0
    era = (p["ER"] * 7) / ip if ip else 0  # HS baseball: 7-inning games
    whip = (p["BB"] + p["H"]) / ip if ip else 0
    k_per_7 = (p["SO"] * 7) / ip if ip else 0
    bb_per_7 = (p["BB"] * 7) / ip if ip else 0
    return ip, era, whip, k_per_7, bb_per_7


def season_summary(players):
    team_bat = {k: 0 for k in ["GP","PA","AB","H","1B","2B","3B","HR","RBI","R","BB","SO","HBP","SF","SB","TB"]}
    team_pit = {k: 0 for k in ["IP_thirds","BF","H","R","ER","BB","SO","W","L","SV","HBP","#P"]}
    rosters = 0
    for p in players:
        b = bat_line(p)
        q = pit_line(p)
        # Only count if the player actually appeared
        if b["PA"] > 0 or q["IP_thirds"] > 0:
            rosters += 1
        for k in team_bat:
            team_bat[k] += b[k]
        for k in team_pit:
            team_pit[k] += q[k]
    return team_bat, team_pit, rosters


def full_name(rec):
    return f"{rec.get('First','').strip()} {rec.get('Last','').strip()}".strip()


def leaders(players, key_fn, min_pa=20, top=5):
    scored = []
    for p in players:
        b = bat_line(p)
        if b["PA"] < min_pa:
            continue
        val = key_fn(b)
        scored.append((val, full_name(p), b))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:top]


def pitch_leaders(players, key_fn, min_ip_thirds=30, top=5, reverse=True):
    scored = []
    for p in players:
        q = pit_line(p)
        if q["IP_thirds"] < min_ip_thirds:
            continue
        val = key_fn(q)
        scored.append((val, full_name(p), q))
    scored.sort(key=lambda x: x[0], reverse=reverse)
    return scored[:top]


def print_header(title):
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)


def main():
    seasons = {}
    for year, fname in FILES.items():
        players, totals, _keys = load_season(STATS_DIR / fname)
        seasons[year] = {"players": players, "totals": totals}

    # --- Team-level season summary ---
    print_header("TEAM SUMMARY BY SEASON — Gig Harbor Varsity Tides")
    print(f"{'Year':<6}{'G':>4}{'AVG':>7}{'OBP':>7}{'SLG':>7}{'OPS':>7}"
          f"{'R':>5}{'HR':>4}{'SB':>5}   "
          f"{'ERA':>6}{'WHIP':>7}{'IP':>8}{'K':>5}{'BB':>5}{'K/7':>7}")
    for year, s in seasons.items():
        tb, tp, _ = season_summary(s["players"])
        avg, obp, slg, ops = calc_batting_rates(tb)
        ip, era, whip, k7, bb7 = calc_pitching_rates(tp)
        games = int(to_num(s["totals"].get("Batting:GP")))
        print(f"{year:<6}{games:>4}"
              f"{fmt_avg(avg):>7}{fmt_avg(obp):>7}{fmt_avg(slg):>7}{fmt_avg(ops):>7}"
              f"{int(tb['R']):>5}{int(tb['HR']):>4}{int(tb['SB']):>5}   "
              f"{era:>6.2f}{whip:>7.2f}"
              f"{ip_from_thirds(tp['IP_thirds']):>8}"
              f"{int(tp['SO']):>5}{int(tp['BB']):>5}{k7:>7.2f}")

    # --- 3-year combined ---
    print_header("3-YEAR COMBINED (2023–2025)")
    combined_b = {k: 0 for k in ["GP","PA","AB","H","1B","2B","3B","HR","RBI","R","BB","SO","HBP","SF","SB","TB"]}
    combined_p = {k: 0 for k in ["IP_thirds","BF","H","R","ER","BB","SO","W","L","SV","HBP","#P"]}
    total_games = 0
    for year, s in seasons.items():
        tb, tp, _ = season_summary(s["players"])
        for k in combined_b: combined_b[k] += tb[k]
        for k in combined_p: combined_p[k] += tp[k]
        total_games += int(to_num(s["totals"].get("Batting:GP")))
    avg, obp, slg, ops = calc_batting_rates(combined_b)
    ip, era, whip, k7, bb7 = calc_pitching_rates(combined_p)
    print(f"Games: {total_games}")
    print(f"Batting: {fmt_avg(avg)}/{fmt_avg(obp)}/{fmt_avg(slg)}  OPS {fmt_avg(ops)}")
    print(f"  H: {int(combined_b['H'])}  2B: {int(combined_b['2B'])}  3B: {int(combined_b['3B'])}  HR: {int(combined_b['HR'])}")
    print(f"  R: {int(combined_b['R'])}  RBI: {int(combined_b['RBI'])}  BB: {int(combined_b['BB'])}  SO: {int(combined_b['SO'])}  SB: {int(combined_b['SB'])}")
    print(f"Pitching: ERA {era:.2f}  WHIP {whip:.2f}  IP {ip_from_thirds(combined_p['IP_thirds'])}")
    print(f"  K: {int(combined_p['SO'])} ({k7:.2f}/7)   BB: {int(combined_p['BB'])} ({bb7:.2f}/7)")
    print(f"  W-L: {int(combined_p['W'])}-{int(combined_p['L'])}   Pitches thrown: {int(combined_p['#P'])}")

    # --- Leaders per season ---
    for year, s in seasons.items():
        print_header(f"{year} SEASON LEADERS")
        print("-- Batting (min 20 PA) --")
        print("Top AVG:")
        for val, name, b in leaders(s["players"], lambda b: (b["H"]/b["AB"]) if b["AB"] else 0, min_pa=20):
            print(f"  {name:<25} {fmt_avg(val)}  ({int(b['H'])}/{int(b['AB'])})")
        print("Top OPS:")
        for val, name, b in leaders(s["players"], lambda b: sum(calc_batting_rates(b)[1:3]) if b["AB"] else 0, min_pa=20):
            a,o,sl,op = calc_batting_rates(b)
            print(f"  {name:<25} {fmt_avg(op)}  (OBP {fmt_avg(o)} / SLG {fmt_avg(sl)})")
        print("Top RBI:")
        for val, name, b in leaders(s["players"], lambda b: b["RBI"], min_pa=20):
            print(f"  {name:<25} {int(val)} RBI")
        print("Top Hits:")
        for val, name, b in leaders(s["players"], lambda b: b["H"], min_pa=20):
            print(f"  {name:<25} {int(val)} H")
        print("Top SB:")
        for val, name, b in leaders(s["players"], lambda b: b["SB"], min_pa=20):
            print(f"  {name:<25} {int(val)} SB")

        print("-- Pitching (min 10 IP) --")
        print("Best ERA:")
        def era_of(q):
            ip = q["IP_thirds"]/3
            return (q["ER"]*7)/ip if ip else 999
        for val, name, q in pitch_leaders(s["players"], era_of, min_ip_thirds=30, reverse=False):
            print(f"  {name:<25} ERA {val:.2f}   IP {ip_from_thirds(q['IP_thirds'])}  K {int(q['SO'])}  BB {int(q['BB'])}")
        print("Most Ks:")
        for val, name, q in pitch_leaders(s["players"], lambda q: q["SO"], min_ip_thirds=30):
            print(f"  {name:<25} {int(val)} K   IP {ip_from_thirds(q['IP_thirds'])}")
        print("Most Wins:")
        for val, name, q in pitch_leaders(s["players"], lambda q: q["W"], min_ip_thirds=15):
            print(f"  {name:<25} {int(val)} W-{int(q['L'])}L   IP {ip_from_thirds(q['IP_thirds'])}")

    # --- Career (multi-year) leaders for players who appeared in multiple seasons ---
    print_header("MULTI-SEASON CAREER TOTALS (2023–2025)")
    career_bat = {}
    career_pit = {}
    years_played = {}
    for year, s in seasons.items():
        for p in s["players"]:
            name = full_name(p)
            if not name:
                continue
            b = bat_line(p)
            q = pit_line(p)
            if b["PA"] > 0 or q["IP_thirds"] > 0:
                years_played.setdefault(name, set()).add(year)
            cb = career_bat.setdefault(name, {k:0 for k in b})
            for k,v in b.items(): cb[k]+=v
            cp = career_pit.setdefault(name, {k:0 for k in q})
            for k,v in q.items(): cp[k]+=v

    # Players who appeared in 2+ seasons
    multi = [(n, yrs) for n, yrs in years_played.items() if len(yrs) >= 2]
    print(f"Players with 2+ varsity seasons: {len(multi)}")
    multi.sort(key=lambda x: career_bat[x[0]]["H"], reverse=True)
    print(f"\n{'Player':<25}{'Yrs':>4}  {'Bat: AVG/OBP/SLG':<22}{'H':>4}{'HR':>4}{'RBI':>5}   {'Pitch: ERA':<12}{'IP':>7}{'K':>5}")
    for name, yrs in multi:
        cb = career_bat[name]; cp = career_pit[name]
        avg, obp, slg, _ = calc_batting_rates(cb)
        if cp["IP_thirds"] > 0:
            ip_val = cp["IP_thirds"]/3
            era = (cp["ER"]*7)/ip_val
            pstr = f"{era:>6.2f}"
            ipstr = ip_from_thirds(cp["IP_thirds"])
            kstr = str(int(cp["SO"]))
        else:
            pstr = "  —   "
            ipstr = "—"
            kstr = "—"
        triple = f"{fmt_avg(avg)}/{fmt_avg(obp)}/{fmt_avg(slg)}"
        print(f"{name:<25}{len(yrs):>4}  {triple:<22}{int(cb['H']):>4}{int(cb['HR']):>4}{int(cb['RBI']):>5}   {pstr:<12}{ipstr:>7}{kstr:>5}")


if __name__ == "__main__":
    main()
