---
name: google-ads-keyword-overlap
description: Use this skill whenever the user wants to detect duplicate or overlapping keywords across ad groups in a Google Ads account that bid against each other (cannibalisatie). Trigger when the user provides a Google Ads keywords export CSV and wants to find conflicts, when the user mentions keyword overlap, keyword cannibalisatie, duplicate keywords, or wants to clean up account structure. The skill produces an Excel file with detected conflicts color-coded by severity, plus an Editor-importable list of suggested pauses.
---

# Google Ads Keyword Overlap Detector (v1.0)

Identifies duplicate or overlapping keywords across ad groups in a Google Ads account that compete against each other for the same impressions. Outputs a color-coded Excel review report and an Editor-importable list of suggested pauses.

## When to use this skill

Trigger conditions:
- User provides a Google Ads keywords export CSV
- User asks to find duplicate / overlapping / cannibalising keywords
- User asks for a "keyword audit" or "account structure cleanup"
- User mentions keywords competing against each other or wasting budget on duplicates

This is a **per-account, one-time tool** — account structure doesn't change frequently, so running this every 3–6 months per client is typical.

## Conflict types detected

| Type | Severity | What it means |
|---|---|---|
| **Exact duplicate** | High (red) | Same keyword + same match type appears in 2+ ad groups |
| **Match type overlap** | Medium (yellow) | Same keyword text in different match types across ad groups (e.g. Broad in AG-A, Exact in AG-B). Broad absorbs Exact's traffic. |
| **Phrase contains** | Medium (yellow) | A Broad keyword is contained as a sub-phrase inside other keywords in different ad groups |
| **Close variant** | Low (green) | Singular/plural / minor stem variations across ad groups (Google merges these automatically since 2018; usually safe to leave) |

## Annotations

The output flags two situations as "probably intentional, review only":

- **Probably geo-segmented** — campaign names suggest different country/region targeting (e.g. `Transport NL` vs `Transport DE`). The skill detects this via campaign-name patterns containing country codes (NL, DE, FR, UK, etc.) or region tokens (EU, EMEA, APAC, Nordic, …) or full country names. Greyed-out in the output, omitted from the "Suggested Pauses" sheet.
- **All paused** — all instances of this conflict are already paused. Informational only; no action needed.

## What the skill does NOT do

- It does not pause anything automatically. Pauses are suggestions for human review.
- It does not pull data from the Google Ads API. Input is a CSV that the user exports manually from Google Ads Editor or the Google Ads UI.
- It does not factor in performance data (cost, conversions). The user decides which instance to keep based on their own knowledge of the account. A future version can take an optional performance CSV as a 2nd input.

## Required CSV columns

Auto-detected (English + Dutch headers supported):

| Required | Aliases |
|---|---|
| Keyword / Zoekwoord | required |
| Match type / Type overeenkomst | required |
| Campaign / Campagne | required |
| Ad group / Advertentiegroep | required |
| Status | required (used to skip paused) |
| Final URL / Uiteindelijke URL | optional |

## How to use

### Step 1 — Get the input CSV

User exports keywords from Google Ads:

**Option A: Google Ads Editor** (recommended for full account dump)
1. Open Editor → connect to the account
2. Keywords view → select all keywords
3. File → Export → Selected keywords → CSV

**Option B: Google Ads UI**
1. Account → Keywords
2. Modify columns: ensure Campaign, Ad group, Keyword, Match type, Status, Final URL are visible
3. Download → CSV

### Step 2 — Run the script

```
python scripts/detect_overlap.py <keywords_csv> [output_path] [--accounts "Acct1,Acct2,..."]
```

Examples:
```bash
# Single-account export
python scripts/detect_overlap.py keywords_export.csv

# MCC export (all sub-accounts)
python scripts/detect_overlap.py mcc_export.csv

# MCC export with sub-account filter (case-insensitive substring match)
python scripts/detect_overlap.py mcc_export.csv \
  --accounts "ARBO,Ewals,Smartgyro,Vetus"
```

### MCC support and account scoping

- If the input CSV has an **Account** column (MCC export), each MCC sub-account is treated as a separate client. Conflict detection runs **per-account**, so a keyword in Smartgyro can never be flagged as conflicting with a keyword in Ewals — they're different clients.
- Use `--accounts "name1,name2"` to limit detection to specific sub-accounts (case-insensitive substring match against the Account column).
- If the CSV has no Account column (single-account export), all rows are treated as one account.

### Step 3 — Report back to the user

After running, summarise:
- Total keywords loaded (active vs paused)
- Total conflicts detected, breakdown by conflict type
- How many are flagged as "probably geo-segmented" (likely intentional)
- Path to the output xlsx
- Note: pauses are suggestions, not commands — user reviews and decides

## Output structure

**Sheet 1: "Overlap Issues"**
Color-coded by severity:
- 🔴 Red — Exact duplicate (highest priority)
- 🟡 Yellow — Match type overlap or Phrase contains
- 🟢 Green — Close variant
- ⚪ Grey — geo-segmented or all-paused (de-emphasised)

Columns: conflict type, keywords, match types, ad groups, campaigns, geo-flag, paused-flag, # instances.

**Sheet 2: "Suggested Pauses"** (Editor-importable)
Paste-ready CSV-like format:
- Campaign, Ad group, Keyword (with `[exact]` / `"phrase"` decoration), Match type, Status="Paused", Conflict type, Reason

For each conflict, the skill suggests pausing all-but-the-first instance. Geo-segmented and close-variant conflicts are excluded from this sheet (typically intentional or low impact).

User selects which rows to apply, copies to Google Ads Editor, sets status accordingly.

**Sheet 3: "Legend"**
Reference for the conflict types and annotations.

## Limitations and caveats

- **Heuristic matching** — close-variant detection uses light stemming, not a full NLP stemmer. Some plural/variant pairs may be missed; some unrelated keywords may be flagged incorrectly. Treat the close-variant tier as "review only".
- **Geo-segmentation detection** uses a hardcoded list of country/region tokens. If campaign names use unconventional naming (e.g. project codes, city names), the heuristic will not flag them. This is a feature, not a bug — better to over-flag conflicts than auto-skip real ones.
- **No performance ranking yet** — the "Suggested Pauses" sheet picks the first row arbitrarily as the "keep" instance. A future version with optional performance data could rank these intelligently. For now, the user decides which to keep.
- **Skips paused keywords** for active conflict detection — paused keywords don't compete for impressions. They may still appear annotated in the output if all instances are paused (then the conflict is shown but marked "all paused" for completeness).

## Versioning

- v1.0 — initial release: 4 conflict types, geo-segmentation heuristic, Editor-importable Suggested Pauses sheet, color-coded review.
