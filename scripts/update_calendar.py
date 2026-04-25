#!/usr/bin/env python3
"""
Replace calendar.html table bodies, FullCalendar event arrays,
and summary stats with newly generated data from script output files.
"""

import re
import os

BASE = "/Users/bsleeter/Documents/Documents - Benjamin's MacBook Pro/GH Baseball"
CALENDAR = os.path.join(BASE, "calendar.html")
SCRIPTS  = os.path.join(BASE, "scripts")

def read(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def write(path, content):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

# --- Load source files ---
html = read(CALENDAR)
lines = html.split("\n")

varsity_table = read(os.path.join(SCRIPTS, "output_varsity_table.html"))
jv_table      = read(os.path.join(SCRIPTS, "output_jv_table.html"))
cteam_table   = read(os.path.join(SCRIPTS, "output_cteam_table.html"))

varsity_events = read(os.path.join(SCRIPTS, "output_varsity_events.js"))
jv_events      = read(os.path.join(SCRIPTS, "output_jv_events.js"))
cteam_events   = read(os.path.join(SCRIPTS, "output_cteam_events.js"))

# --- Identify <tbody> blocks by line numbers (1-indexed in grep, 0-indexed here) ---
# Find all <tbody> and </tbody> lines
tbody_opens = []
tbody_closes = []
for i, line in enumerate(lines):
    stripped = line.strip()
    if stripped == "<tbody>":
        tbody_opens.append(i)
    elif stripped == "</tbody>":
        tbody_closes.append(i)

print(f"Found <tbody> at lines (0-indexed): {tbody_opens}")
print(f"Found </tbody> at lines (0-indexed): {tbody_closes}")

assert len(tbody_opens) == 3, f"Expected 3 <tbody>, found {len(tbody_opens)}"
assert len(tbody_closes) == 3, f"Expected 3 </tbody>, found {len(tbody_closes)}"

# Build new lines array by replacing tbody content
# We'll work backwards so line numbers don't shift
replacements = [
    (tbody_opens[2], tbody_closes[2], cteam_table),   # C Team
    (tbody_opens[1], tbody_closes[1], jv_table),       # JV
    (tbody_opens[0], tbody_closes[0], varsity_table),  # Varsity
]

for open_idx, close_idx, new_content in replacements:
    # Keep the <tbody> line and </tbody> line, replace everything in between
    # Ensure new_content doesn't have trailing newline that would cause blank line
    new_content = new_content.rstrip("\n")
    new_lines_segment = [lines[open_idx]] + [new_content] + [lines[close_idx]]
    lines[open_idx:close_idx+1] = new_lines_segment

# Rejoin for regex replacements on the event arrays and stats
html = "\n".join(lines)

# --- Replace FullCalendar event arrays ---
# Pattern: var varsityEvents = [...];
def replace_events(html, var_name, new_events_content):
    # new_events_content is the raw JS lines (no surrounding [ ])
    new_events_content = new_events_content.strip()
    # Match var <name> = [...]; across multiple lines
    pattern = r'(var\s+' + var_name + r'\s*=\s*\[).*?(\];)'
    replacement = r'\g<1>\n' + new_events_content + r'\n\2'
    new_html, count = re.subn(pattern, replacement, html, flags=re.DOTALL)
    assert count == 1, f"Expected 1 replacement for {var_name}, got {count}"
    return new_html

html = replace_events(html, "varsityEvents", varsity_events)
html = replace_events(html, "jvEvents", jv_events)
html = replace_events(html, "cteamEvents", cteam_events)

# --- Update summary stats ---
# Stats appear in order: Varsity, JV, C Team
# Each block has: Total Events, Home Games, Away Games, Practices
stats = [
    # (total, home, away, practices)
    (58, 11, 10, 33),  # Varsity
    (51, 11, 9, 28),   # JV
    (57, 6, 13, 35),   # C Team
]

stat_labels = ["Total Events", "Home Games", "Away Games", "Practices"]

for team_idx, (total, home, away, practices) in enumerate(stats):
    values = [total, home, away, practices]
    for label, value in zip(stat_labels, values):
        # Find the nth occurrence (team_idx-th)
        # Pattern: <div class="num">XX</div><div class="lbl">Label</div>
        pattern = r'(<div class="num">)\d+(</div><div class="lbl">' + re.escape(label) + r'</div>)'
        
        # Find all matches and replace the correct one
        matches = list(re.finditer(pattern, html))
        assert len(matches) >= team_idx + 1, f"Not enough matches for {label}"
        
        match = matches[team_idx]
        html = html[:match.start()] + match.group(1) + str(value) + match.group(2) + html[match.end():]

# --- Write output ---
write(CALENDAR, html)

new_line_count = html.count("\n") + 1
print(f"\nDone! calendar.html written with {new_line_count} lines.")
