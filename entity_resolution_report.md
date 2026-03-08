# Entity Resolution and Master Dataset Construction

## Inputs
- `people_dataset_15000_challenges_with_overlap.xlsx` (people/contact dataset, 15,000 rows).
- `fake_vehicle_people_dataset_15000_challenges.xlsx` (vehicle/person dataset, 15,000 rows).

## What was done
1. Built a pure-Python XLSX reader (ZIP/XML parser) to ingest both workbooks without external dependencies.
2. Standardized matching fields:
   - Names: lowercased, stripped of non-letters.
   - Phones: reduced to digits, normalized to up to 11 digits.
   - Dates: parsed from mixed formats into `YYYY-MM-DD`.
   - Address text: normalized abbreviations and punctuation.
3. Generated candidate pairs with blocking keys:
   - Exact phone match.
   - `(last_name, first_initial)`.
   - `(last_name, city)`.
4. Scored each candidate with weighted features (name, phone, date, city/state, street similarity).
5. Performed greedy one-to-one assignment on highest score pairs.
6. Created:
   - `matched_pairs.csv` with cross-dataset links and evidence fields.
   - `master_people_dataset.csv` with consolidated person entities and source coverage labels.
   - `resolution_summary.txt` with headline counts.

## Outcome
- Total records in A: **15,000**
- Total records in B: **15,000**
- Cross-dataset matched entities: **662**
- Master dataset entities (A union B with overlap merged): **29,338**

## Output files
- `entity_resolution.py`
- `matched_pairs.csv`
- `master_people_dataset.csv`
- `resolution_summary.txt`
