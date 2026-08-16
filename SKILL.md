---
name: trip-planner
description: >-
  Professional travel-planning assistant. Turns a destination, number of days,
  companions, budget, and preferences into a geographically clustered itinerary
  (one district per day, single-direction flow, 20% slack) and a formatted Excel
  workbook (.xlsx) with "Daily Itinerary / Budget / Checklist" sheets. Trigger
  on requests like "plan a trip to X", "make me a 5-day itinerary", "旅行行程规划",
  "做一份 X 天攻略", "itinerary / trip plan", "行程表 Excel", "目的地 + 天数 + 预算",
  or any ask for a ready-to-execute travel plan — even when the user does not
  explicitly say "generate an Excel", produce a real deliverable whenever they
  want a trip they can actually follow.
---

# Trip Planner

Turn a vague travel request into an executable, printable plan. Two phases with
strict separation of concerns:

1. **Plan (model)**: geographic clustering + pacing + single-direction flow →
   structured itinerary data.
2. **Render (script)**: `scripts/build_trip_excel.py` renders that data into a
   consistently styled `.xlsx`.

The script owns only deterministic layout (colors, borders, column widths,
freeze panes, budget SUM). It makes **zero** "where to go / how to route"
decisions. Every run produces an identical look without re-deriving openpyxl
styling.

## When to use

Trigger on travel-planning intent even if "Excel" is never mentioned. The
deliverable is a plan, and the spreadsheet is just its format.

## Workflow

### 1. Extract requirements

Pull these fields from the request; fill gaps with sensible defaults and call
out which values were defaulted when reporting back.

| Field | Meaning | When missing |
| --- | --- | --- |
| `destination` | city / region | **Required** — ask via AskUserQuestion |
| `days` | total trip days | **Required** — ask |
| companions | count, age, stamina, relationship | default "1–2 adults, average stamina" |
| budget | shoestring / comfort / luxury or a number | default "mid-range comfort" |
| special needs | food, kids, photo, offbeat, summer-avoid, etc. | default "no strong preference" |

Destination and days are the foundation — never guess them.

### 2. Plan the itinerary (golden rules)

Three rules; break them and the trip becomes "backtracking + rushing":

1. **POI clustering**: each day is confined to **one** district — no cross-city
   backtracking. Cluster attractions on a map first, then assign per day.
2. **Pacing**: one anchor attraction + 1–2 light stopovers per day, with 20%
   slack so a day is never packed.
3. **Single-direction flow**: `hotel → breakfast → morning sights → lunch →
   afternoon indoor/shaded → evening/night view → dinner back` — no loops, no
   doubling back.

Each day needs: time slot, activity, transit mode + duration, dwell time,
signature food, and booking/avoid-trap reminders.

### 3. Write the trip data (JSON)

Serialize the plan to `trip_data.json` (schema in the script docstring, or the
condensed version below). This is the **only** place you write data — no layout
logic here. The script validates the schema and fails loudly on mistakes.

```json
{
  "destination": "Osaka / 大阪",
  "days": 5,
  "currency": "¥",
  "dates": "2026-08-01 ~ 2026-08-05",
  "travelers": "2 adults",
  "budget_style": "Mid-range / comfortable",
  "notes": "Bring a portable charger",
  "daily_itinerary": [
    { "day": 1, "slots": [
      { "slot": "08:30-09:00", "activity": "Leave hotel", "duration": "30min",
        "transit": "walk", "dining": "Breakfast: Kuromon Market",
        "tips": "Some stalls closed Monday" }
    ]}
  ],
  "budget": [
    { "category": "Transport", "item": "Round-trip flight", "estimated": 3000,
      "actual": null, "notes": "tax included" }
  ],
  "checklist": [
    { "category": "Documents", "item": "Passport/ID", "status": "", "notes": "check expiry" }
  ]
}
```

Notes: `estimated`/`actual` may be a number (summed) or a string like
"3000/person" (not summed). Empty `status` auto-fills "⬜ 待准备". Budget must
cover transport, lodging, tickets, dining, and contingency; the checklist must
cover documents, clothing, medicine, electronics, and must-have apps.

### 4. Run the script

1. Confirm the dependency: `pip install openpyxl -q` (the script prints the same
   hint if the import fails).
2. Run: `python <skill dir>/scripts/build_trip_excel.py trip_data.json [-o OUT_DIR]`
3. The script prints the **absolute path** of the output; relay it verbatim.

Output name: `<destination>_Trip_Plan_<days>D.xlsx` (destination is filename-sanitized).

### 5. Report back

Print a Markdown overview (one line per day + budget total + key reminders)
first, then the `.xlsx` absolute path.

## Output file

One `.xlsx` workbook:

- **Overview** (optional, only when `dates`/`travelers`/`budget_style`/`notes`
  are present) — key facts at a glance.
- **Daily Itinerary** — day / slot / activity / dwell time / transit / dining /
  tips.
- **Budget** — category / item / estimated / actual / notes, with a live `SUM`
  total row and currency formatting.
- **Checklist** — category / item / status / notes.

Every sheet: deep-blue header (`#1F4E78` + white bold), auto column widths,
frozen header row, thin gray borders, and wrapped text with content-aware row
heights.
