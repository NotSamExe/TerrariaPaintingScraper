# Terraria Paintings Scraper

Made this for my own Terraria world with friends, as we were making it a point to collect every painting in the game. Scrapes all painting data from the [Terraria wiki](https://terraria.wiki.gg/wiki/Paintings) and generates a tracking spreadsheet so you can keep track of which paintings you've collected.

---

## What it produces

| File | Description |
|------|-------------|
| `paintings.xlsx` | Excel workbook with all 193 paintings + tracker |
| `painting_images/` | All painting and placed-preview PNGs downloaded from the wiki |

---

## How to run

Double-click **`Run Scraper.bat`**.

- First run takes like 10 seconds to download all images from the wiki.
- Subsequent runs are faster — already-downloaded images are skipped.
- The Excel file opens automatically when done.

> If Python is not installed, download it from https://python.org (check "Add to PATH" during install), then run the bat again.

---

## How to use the spreadsheet

### Marking a painting as obtained

1. Click the cell in the **Obtained? (T/F)** column (column A) for that painting.
2. Type **`T`** and press Enter.
3. The row turns **green** automatically.


### Columns

| Column | Contents |
|--------|----------|
| Obtained? (T/F) | Your progress — type T or F |
| Painting | Item sprite thumbnail |
| Name | Painting name |
| Size | Tile dimensions (W × H) |
| Source / How to Get | Which NPC or location it comes from |
| Description/Details | Full description from the wiki |
| Buy Price | Cost from an NPC (if sold) |
| Sell Price | Value when sold to an NPC |
| Tooltip / Artist | In-game tooltip and artist credit |
| Placed Preview | How it looks placed in the world |

---



## Re-running after a wiki update

Just double-click `Run Scraper.bat` again. The spreadsheet is fully regenerated from the latest wiki data.

> **Note:** Re-running overwrites `paintings.xlsx` and resets all your T/F entries. If you want to keep your progress, copy your T/F column somewhere before re-running.
