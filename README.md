# Google Ads Keyword Overlap Detector

Detects duplicate or overlapping keywords across ad groups in a single Google Ads account that bid against each other (cannibalisatie). Outputs a color-coded Excel review report and an Editor-importable list of suggested pauses.

## What it detects

| Conflict type | Severity | Meaning |
|---|---|---|
| Exact duplicate | High | Same keyword + same match type in 2+ ad groups |
| Match type overlap | Medium | Same keyword text, different match types across ad groups |
| Phrase contains | Medium | Broad keyword X contained as sub-phrase in other ad groups' keywords |
| Close variant | Low | Singular/plural / minor stem variations across ad groups |

## Annotations

- **Probably geo-segmented** — campaigns ending in different country/region tokens (NL/DE/EU/etc.) are flagged as likely-intentional and excluded from the suggested-pause list.
- **All paused** — informational; conflict already inactive.

## Usage

```bash
python scripts/detect_overlap.py <keywords_csv> [output_path]
```

Input: keywords export CSV from Google Ads Editor or Google Ads UI. Required columns: Keyword, Match type, Campaign, Ad group, Status. English and Dutch headers both work.

Output: Excel with 3 sheets — Overlap Issues (color-coded), Suggested Pauses (Editor-importable), Legend.

## Files

- `scripts/detect_overlap.py` — main script, single-file Python with no external state
- `SKILL.md` — instructions for the AI when invoked
- `README.md` — this file

## Limitations

- Detection is heuristic; close-variant matching uses light stemming.
- Geo-segmentation detection uses a hardcoded list of country/region tokens.
- No performance ranking yet — the user picks which conflicts to actually pause.
- Skips paused keywords from active conflict detection.

## Future ideas

- Optional 2nd input: performance CSV → smarter "keep best performer" recommendations
- Cross-account overlap detection (requires API access — TODO when API approved)
- Negative keyword conflict detection (negative blocking positive in same account)
