# Data Quality Notes — athlete_events.csv
**Generated:** 2026-03-01 12:06

## Dataset Overview
- Rows: 271,116
- Columns: 15
- Year range: 1896 – 2016
- Unique athletes: 135,571
- Unique NOCs: 230
- Unique sports: 66
- Unique events: 765

## Missing Data
- `Medal`: 231,333.0 missing (85.33%)
- `Weight`: 62,875.0 missing (23.19%)
- `Height`: 60,171.0 missing (22.19%)
- `Age`: 9,474.0 missing (3.49%)

## Anomalies
- Age outliers (<10 or >80): 8 rows
- Height outliers (<130 or >230 cm): 8 rows
- Weight outliers (<30 or >200 kg): 22 rows

## Recommended Cleaning Steps
1. Impute or flag missing Age/Height/Weight by sport median
2. Standardize Team/NOC names to ISO 3166-1 alpha-3
3. Drop duplicate rows if confirmed non-meaningful
4. Create `Medal_Won` binary flag (1 = medalist, 0 = non-medalist)
5. Filter to Summer Olympics for LA 2028 relevance (retain Winter for benchmarking)