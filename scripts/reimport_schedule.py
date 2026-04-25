#!/usr/bin/env python3
"""Reimport schedule data from the Master Schedule Excel file."""
import openpyxl
import re
import json
from datetime import datetime, timedelta

wb = openpyxl.load_workbook('Schedule/2026 Master Schedule.xlsx', data_only=True)
ws = wb['Sheet1 (2)']

# Read all rows (skip header rows)
rows = list(ws.iter_rows(values_only=True))
header1 = rows[0]  # (None, 'Varsity', None, None, 'JV', None, None, 'C Team', ...)
header2 = rows[1]  # (Date, Type, Time, Location, Type, Time, Location, Type, Time, Location, ...)

data_rows = rows[2:]

def parse_date(date_str):
    """Parse date string like 'Monday, March 2, 2026' to datetime."""
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str.strip(), '%A, %B %d, %Y')
    except:
        return None

def parse_time_to_24h(time_str):
    """Parse time like '4:30pm', '7pm', '2:30-5:00 pm' and return start/end in HH:MM format."""
    if not time_str:
        return None, None
    time_str = time_str.strip().replace('\xa0', ' ')

    # Handle range like "2:30-5:00 pm" or "5:00 - 9:00 pm" or "6:30 pm - 9:00 pm" or "7:00 pm - 9:00 pm"
    range_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(?:pm|am)?\s*-\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if range_match:
        start_str = range_match.group(1)
        end_str = range_match.group(2)
        meridiem = range_match.group(3).lower()

        start_24 = convert_to_24h(start_str, meridiem)
        end_24 = convert_to_24h(end_str, meridiem)

        # If start > end, start is AM
        if start_24 and end_24:
            sh, sm = map(int, start_24.split(':'))
            eh, em = map(int, end_24.split(':'))
            if sh > eh:
                # start must be AM
                start_24 = convert_to_24h(start_str, 'am')

        return start_24, end_24

    # Handle "6:30 pm - 9:00 pm" style
    range_match2 = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)\s*-\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if range_match2:
        start_str = range_match2.group(1)
        start_mer = range_match2.group(2).lower()
        end_str = range_match2.group(3)
        end_mer = range_match2.group(4).lower()
        return convert_to_24h(start_str, start_mer), convert_to_24h(end_str, end_mer)

    # Handle single time like "4pm", "4:30pm", "12pm"
    single_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if single_match:
        t = convert_to_24h(single_match.group(1), single_match.group(2).lower())
        return t, None

    # Handle comma-separated times like "11am, 2pm" or "11am, 1pm"
    multi_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)\s*,\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if multi_match:
        t1 = convert_to_24h(multi_match.group(1), multi_match.group(2).lower())
        t2 = convert_to_24h(multi_match.group(3), multi_match.group(4).lower())
        return t1, t2  # Return both game times

    return None, None

def convert_to_24h(time_str, meridiem):
    """Convert time like '4:30' with meridiem 'pm' to '16:30'."""
    parts = time_str.split(':')
    hour = int(parts[0])
    minute = int(parts[1]) if len(parts) > 1 else 0

    if meridiem == 'pm' and hour != 12:
        hour += 12
    elif meridiem == 'am' and hour == 12:
        hour = 0

    return f'{hour:02d}:{minute:02d}'

def format_time_display(time_str):
    """Format time for display like '4:00 PM', '7:00 - 9:00 PM'."""
    if not time_str:
        return 'TBD'
    time_str = time_str.strip().replace('\xa0', ' ')

    # Range time like "2:30-5:00 pm" or "8:00 - 3:00 pm"
    range_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(?:pm|am)?\s*-\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if range_match:
        start_h = int(range_match.group(1).split(':')[0])
        end_h = int(range_match.group(2).split(':')[0])
        end_mer = range_match.group(3).upper().strip()
        start_mer = end_mer
        # Only flip to AM when start > end numerically (e.g. 8 > 3 means 8am-3pm)
        if end_mer == 'PM' and start_h > end_h:
            start_mer = 'AM'
        s = format_single_time(range_match.group(1), start_mer)
        e = format_single_time(range_match.group(2), end_mer)
        return f'{s} - {e}'

    range_match2 = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)\s*-\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if range_match2:
        s = format_single_time(range_match2.group(1), range_match2.group(2))
        e = format_single_time(range_match2.group(3), range_match2.group(4))
        return f'{s} - {e}'

    # Single time
    single_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if single_match:
        return format_single_time(single_match.group(1), single_match.group(2))

    # Multi time like "11am, 1pm"
    multi_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)\s*,\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', time_str, re.IGNORECASE)
    if multi_match:
        t1 = format_single_time(multi_match.group(1), multi_match.group(2))
        t2 = format_single_time(multi_match.group(3), multi_match.group(4))
        return f'{t1} & {t2}'

    return time_str.strip()

def format_single_time(time_str, meridiem):
    """Format a single time for display."""
    parts = time_str.split(':')
    hour = int(parts[0])
    minute = int(parts[1]) if len(parts) > 1 else 0
    m = meridiem.upper().strip()
    if minute == 0:
        return f'{hour}:{minute:02d} {m}'
    return f'{hour}:{minute:02d} {m}'

def classify_event(type_str):
    """Classify event type from spreadsheet."""
    if not type_str:
        return None
    t = type_str.strip().lower()
    if 'tryout' in t:
        return 'Tryouts'
    elif 'jamboree' in t:
        return 'Jamboree'
    elif 'game' in t:
        if '@' in t or 'at ' in t.lower():
            return 'Away Game'
        elif 'vs' in t:
            return 'Home Game'
        else:
            return 'Game'
    elif 'practice' in t:
        return 'Practice'
    return None

def extract_opponent(type_str):
    """Extract opponent name from type string like 'Game vs Kelso' or 'Game @ Rogers'."""
    if not type_str:
        return ''
    # Remove 'Game' prefix
    t = re.sub(r'^Game\s+', '', type_str.strip(), flags=re.IGNORECASE)
    # Handle "@ Rogers" or "at Rogers"
    m = re.match(r'(?:@|at)\s+(.+)', t, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Handle "vs Rogers" or "vs. Rogers"
    m = re.match(r'vs\.?\s+(.+)', t, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return t.strip()

# Process each team's data
teams = {
    'varsity': {'col_offset': 1, 'events': [], 'table_rows': []},
    'jv': {'col_offset': 4, 'events': [], 'table_rows': []},
    'cteam': {'col_offset': 7, 'events': [], 'table_rows': []},
}

WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# FullCalendar color schemes
FC_COLORS = {
    'Practice': {'bg': '#E3F2FD', 'border': '#90CAF9', 'text': '#1565C0'},
    'Tryouts': {'bg': '#F3E5F5', 'border': '#7B1FA2', 'text': '#7B1FA2'},
    'Jamboree': {'bg': '#FFF3E0', 'border': '#E65100', 'text': '#E65100'},
    'Home Game': {'bg': '#E8F5E9', 'border': '#2E7D32', 'text': '#2E7D32'},
    'Away Game': {'bg': '#BBDEFB', 'border': '#1565C0', 'text': '#1565C0'},
}

for row in data_rows:
    date = parse_date(row[0])
    if not date:
        continue

    weekday = WEEKDAYS[date.weekday()]
    month = MONTHS[date.month - 1]
    day = date.day
    date_str = date.strftime('%Y-%m-%d')
    date_display = f'{weekday}, {month} {day}'

    for team_name, team_info in teams.items():
        offset = team_info['col_offset']
        evt_type = row[offset] if offset < len(row) else None
        evt_time = row[offset + 1] if offset + 1 < len(row) else None
        evt_loc = row[offset + 2] if offset + 2 < len(row) else None

        if not evt_type:
            continue

        evt_type_str = str(evt_type).strip() if evt_type else ''
        evt_time_str = str(evt_time).strip() if evt_time else ''
        evt_loc_str = str(evt_loc).strip() if evt_loc else ''

        category = classify_event(evt_type_str)
        if not category:
            continue

        opponent = extract_opponent(evt_type_str)
        time_display = format_time_display(evt_time_str)
        start_24, end_24 = parse_time_to_24h(evt_time_str)

        # === Calendar Table Row ===
        is_game = category in ('Home Game', 'Away Game')
        tr_class = ' class="event-game"' if is_game else ''

        if category == 'Practice':
            badge = '<span class="badge badge-practice">Practice</span>'
        elif category == 'Tryouts':
            badge = '<span class="badge badge-tryout">Tryouts</span>'
        elif category == 'Jamboree':
            badge = '<span class="badge badge-roundrobin">Jamboree</span> <span style="font-size:0.8rem; color:var(--text-light);">vs. Peninsula &amp; Silas (3-inn games)</span>'
        elif category == 'Home Game':
            badge = f'<span class="badge badge-home">Home</span> vs {opponent}'
        elif category == 'Away Game':
            badge = f'<span class="badge badge-away">Away</span> @ {opponent}'
        else:
            badge = evt_type_str

        table_row = f'''                <tr{tr_class}>
                    <td class="date-cell"><div class="date-day">{date_display}</div></td>
                    <td class="time-cell">{time_display}</td>
                    <td>{evt_loc_str}</td>
                    <td>{badge}</td>
                </tr>'''
        team_info['table_rows'].append(table_row)

        # === FullCalendar Events ===
        colors = FC_COLORS.get(category, FC_COLORS['Practice'])

        # Handle doubleheader (comma-separated times like "11am, 1pm" or "11am, 2pm")
        multi_match = re.match(r'(\d{1,2}(?::\d{2})?)\s*(pm|am)\s*,\s*(\d{1,2}(?::\d{2})?)\s*(pm|am)', evt_time_str, re.IGNORECASE) if evt_time_str else None

        if multi_match and is_game:
            # Doubleheader - create two events
            t1 = convert_to_24h(multi_match.group(1), multi_match.group(2).lower())
            t2 = convert_to_24h(multi_match.group(3), multi_match.group(4).lower())

            prefix = '@ ' if category == 'Away Game' else 'vs '

            # Game 1
            g1_end_h = int(t1.split(':')[0]) + 2
            g1_end_m = t1.split(':')[1]
            g1_end = f'{g1_end_h:02d}:{g1_end_m}'

            evt1 = f"{{title:'{prefix}{opponent} (G1)',start:'{date_str}T{t1}:00',end:'{date_str}T{g1_end}:00'," \
                   f"backgroundColor:'{colors['bg']}',borderColor:'{colors['border']}',textColor:'{colors['text']}'," \
                   f"extendedProps:{{type:'{category}',location:'{evt_loc_str}'}}}}"
            team_info['events'].append(evt1)

            # Game 2
            g2_end_h = int(t2.split(':')[0]) + 2
            g2_end_m = t2.split(':')[1]
            g2_end = f'{g2_end_h:02d}:{g2_end_m}'

            evt2 = f"{{title:'{prefix}{opponent} (G2)',start:'{date_str}T{t2}:00',end:'{date_str}T{g2_end}:00'," \
                   f"backgroundColor:'{colors['bg']}',borderColor:'{colors['border']}',textColor:'{colors['text']}'," \
                   f"extendedProps:{{type:'{category}',location:'{evt_loc_str}'}}}}"
            team_info['events'].append(evt2)
        elif start_24:
            # Single event
            if is_game:
                prefix = '@ ' if category == 'Away Game' else 'vs '
                title = f'{prefix}{opponent}'
                # Games: add 2.5 hour duration
                end_h = int(start_24.split(':')[0]) + 2
                end_m = int(start_24.split(':')[1]) + 30
                if end_m >= 60:
                    end_h += 1
                    end_m -= 60
                end_time = f'{end_h:02d}:{end_m:02d}'
            elif category == 'Jamboree':
                title = 'Jamboree'
                end_time = end_24 if end_24 else '15:00'
            elif category == 'Tryouts':
                title = 'Tryouts'
                end_time = end_24 if end_24 else '21:00'
            else:
                title = 'Practice'
                end_time = end_24 if end_24 else start_24

            evt_str = f"{{title:'{title}',start:'{date_str}T{start_24}:00',end:'{date_str}T{end_time}:00'," \
                      f"backgroundColor:'{colors['bg']}',borderColor:'{colors['border']}',textColor:'{colors['text']}'," \
                      f"extendedProps:{{type:'{category}',location:'{evt_loc_str}'}}}}"
            team_info['events'].append(evt_str)
        elif evt_time_str.upper() == 'TBD' or not evt_time_str:
            # All-day event
            if is_game:
                prefix = '@ ' if category == 'Away Game' else 'vs '
                title = f'{prefix}{opponent}'
            else:
                title = category

            evt_str = f"{{title:'{title}',start:'{date_str}',allDay:true," \
                      f"backgroundColor:'{colors['bg']}',borderColor:'{colors['border']}',textColor:'{colors['text']}'," \
                      f"extendedProps:{{type:'{category}',location:'{evt_loc_str}'}}}}"
            team_info['events'].append(evt_str)

# Count stats
for team_name, team_info in teams.items():
    home_games = sum(1 for e in team_info['events'] if "'Home Game'" in e)
    away_games = sum(1 for e in team_info['events'] if "'Away Game'" in e)
    practices = sum(1 for e in team_info['events'] if "'Practice'" in e)
    tryouts = sum(1 for e in team_info['events'] if "'Tryouts'" in e)
    jamboree = sum(1 for e in team_info['events'] if "'Jamboree'" in e)
    total = len(team_info['events'])

    print(f"\n=== {team_name.upper()} ===")
    print(f"Total events: {total}")
    print(f"Home games: {home_games}")
    print(f"Away games: {away_games}")
    print(f"Practices: {practices}")
    print(f"Tryouts: {tryouts}")
    if jamboree:
        print(f"Jamboree: {jamboree}")

# Output calendar table rows
for team_name in ['varsity', 'jv', 'cteam']:
    filename = f'scripts/output_{team_name}_table.html'
    with open(filename, 'w') as f:
        f.write('\n'.join(teams[team_name]['table_rows']))
    print(f"\nWrote {filename}")

# Output FullCalendar events
for team_name in ['varsity', 'jv', 'cteam']:
    filename = f'scripts/output_{team_name}_events.js'
    with open(filename, 'w') as f:
        f.write(',\n'.join(teams[team_name]['events']))
    print(f"Wrote {filename}")

# Also output schedule.html game data for comparison
print("\n\n=== SCHEDULE.HTML GAME DATA ===")
for team_name in ['varsity', 'jv', 'cteam']:
    print(f"\n--- {team_name.upper()} GAMES ---")
    for row in data_rows:
        date = parse_date(row[0])
        if not date:
            continue
        offset = teams[team_name]['col_offset']
        evt_type = row[offset] if offset < len(row) else None
        evt_time = row[offset + 1] if offset + 1 < len(row) else None
        evt_loc = row[offset + 2] if offset + 2 < len(row) else None

        if not evt_type:
            continue

        category = classify_event(str(evt_type))
        if category in ('Home Game', 'Away Game'):
            opponent = extract_opponent(str(evt_type))
            time_display = format_time_display(str(evt_time)) if evt_time else 'TBD'
            badge_type = 'home' if category == 'Home Game' else 'away'
            print(f"  {date.strftime('%b %d')} | {category} | {opponent} | {time_display} | {evt_loc}")
