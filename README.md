# Trip Planner · Claude Code Skill

> Turn a vague travel request into a **print-ready, executable itinerary** — a
> geographically clustered plan that actually flows, rendered into a polished
> multi-sheet Excel workbook.

A [Claude Code](https://claude.com/claude-code) **skill**: a small, installable
package that teaches Claude a two-phase workflow — *plan* the trip (routing,
pacing, logistics), then *render* it with a deterministic script that produces
a consistently styled `.xlsx` every time.

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-green.svg)](https://www.python.org/)

---

## What you get

One `.xlsx` workbook with up to four sheets:

| Sheet | Contents |
| --- | --- |
| **Overview** *(optional)* | Destination, dates, travelers, budget style, notes |
| **Daily Itinerary** | Day / time slot / activity / dwell time / transit / dining / tips |
| **Budget** | Category / item / estimated / actual / notes — with a live `SUM` total |
| **Checklist** | Category / item / packing status / notes |

Every sheet ships with the same visual language: deep-blue header (`#1F4E78`),
auto column widths, a frozen header row, thin gray borders, and wrapped text
with content-aware row heights.

### The plan is the product

The formatting is deterministic, but the *planning* follows three hard rules
that separate a real itinerary from a "backtracking + rushing" list:

1. **POI clustering** — each day is confined to one district. No cross-city
   zig-zagging.
2. **Pacing** — one anchor attraction + 1–2 light stopovers, with **20% slack**
   so a day is never packed.
3. **Single-direction flow** — `hotel → breakfast → morning → lunch → afternoon
   → evening → dinner back`, never doubling back.

---

## Installation

Skills are directories under `~/.claude/skills/`. Install by cloning:

```bash
git clone https://github.com/<your-username>/trip-planner.git \
  ~/.claude/skills/trip-planner
```

Then install the one runtime dependency:

```bash
pip install openpyxl
```

That's it. Restart Claude Code (or start a new session) and the skill is
available.

---

## Usage

Just ask for a trip plan. The skill triggers on travel-planning intent even
when you never mention "Excel" — the deliverable is the plan, the spreadsheet
is only its format.

**English**

> Plan a 5-day trip to Osaka for two adults, mid-range budget, focused on food
> and a couple of easy day trips.

**中文**

> 帮我规划一份大阪 5 天攻略，两个人，中档预算，想吃美食，再排一两个轻松的一日游。

Claude then:

1. Extracts destination, days, companions, budget, and preferences (asking when
   destination or days are missing).
2. Plans the clustered, paced, single-direction itinerary.
3. Writes a `trip_data.json` and runs `scripts/build_trip_excel.py`.
4. Reports a Markdown summary plus the absolute path to the `.xlsx`.

The output file is named `<destination>_Trip_Plan_<days>D.xlsx`.

---

## How it works

```
┌─────────────────────────────────────────────────────────────┐
│  Model (planning)          Script (rendering)               │
│  ─────────────────          ──────────────────              │
│  · cluster POIs            · sheet layout & styling         │
│  · pace & slack            · auto column widths             │
│  · single-direction flow   · freeze panes & borders         │
│  · budget / checklist      · SUM total + currency format    │
│            │                          │                      │
│            └──────────┬───────────────┘                      │
│                       ▼                                      │
│                 trip_data.json  ──►  formatted .xlsx         │
└─────────────────────────────────────────────────────────────┘
```

The script makes **zero** travel decisions — it only formats what the model
plans. This separation keeps every run visually consistent and means the
planner never re-derives openpyxl styling.

### Input schema

The script consumes a single JSON file (see
[`examples/trip_data.json`](examples/trip_data.json) for a complete, realistic
example):

```jsonc
{
  "destination": "Osaka / 大阪",
  "days": 5,
  "currency": "¥",                            // optional; "" disables symbol
  "dates": "2026-08-01 ~ 2026-08-05",         // optional → Overview sheet
  "travelers": "2 adults",                    // optional
  "budget_style": "Mid-range / comfortable",  // optional
  "notes": "Bring a portable charger",        // optional
  "daily_itinerary": [
    { "day": 1, "slots": [
      { "slot": "08:30-09:00", "activity": "Leave hotel", "duration": "30min",
        "transit": "walk", "dining": "Breakfast: Kuromon Market",
        "tips": "Some stalls closed Monday" }
    ]}
  ],
  "budget": [
    { "category": "Transport", "item": "Round-trip flight",
      "estimated": 3000, "actual": null, "notes": "tax included" }
  ],
  "checklist": [
    { "category": "Documents", "item": "Passport/ID",
      "status": "", "notes": "check expiry" }
  ]
}
```

- `estimated` / `actual` may be a **number** (summed into the total) or a
  **string** like `"3000/person"` (rendered as text, not summed).
- An empty `status` auto-fills `⬜ 待准备`.
- The `Overview` sheet appears only when at least one of the optional text
  fields is present.
- Input is **validated** — the script exits with a numbered list of problems
  instead of silently producing an empty workbook.

### CLI reference

```bash
python scripts/build_trip_excel.py trip_data.json [-o OUT_DIR] [--version]
```

| Arg | Meaning |
| --- | --- |
| `trip_data` | path to the input JSON |
| `-o, --out-dir` | output directory (default: current dir) |
| `--version` | print the script version |

On success it prints the absolute path of the generated workbook to stdout.

---

## Development

Run the test suite (requires `pytest`):

```bash
pip install -r requirements.txt pytest
pytest -q
```

The tests cover input validation, the formatting helpers, sheet structure, the
budget `SUM` formula, and an end-to-end CLI run against the example data.

### Project layout

```
trip-planner/
├── SKILL.md                  # the skill definition Claude Code loads
├── scripts/
│   └── build_trip_excel.py   # deterministic .xlsx renderer
├── examples/
│   └── trip_data.json        # complete, realistic example input
├── tests/
│   └── test_build_trip_excel.py
├── requirements.txt
└── LICENSE
```

---

## License

[MIT](LICENSE) © Josh Huang
