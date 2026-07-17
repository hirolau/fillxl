# Excel Test File: test_template_5_wide_table.xlsx — Sheet Descriptions

A single sheet with a 20-column-wide city-data table used for benchmarking and wide-table testing.

---

## Sheet: CityData

**Table:** `CityData` — `A1:T11`

10 rows of city data with columns covering strings, numbers, booleans, dates, times, and mixed types.

| Column              | Type        | Example              |
|---------------------|-------------|----------------------|
| City                | string      | Tokyo                |
| Country             | string      | Japan                |
| Population          | int         | 13,960,000           |
| Area (km²)          | float       | 2194.07              |
| Founded             | string      | 1457-01-01           |
| Is Capital          | bool        | TRUE                 |
| Elevation (m)       | int         | 40                   |
| Avg Temp (°C)       | float       | 15.8                 |
| Timezone            | string      | UTC+9                |
| Last Census         | datetime    | 2020-10-01           |
| Has Metro           | bool        | TRUE                 |
| GDP (Billion $)     | float       | 1920.5               |
| Airports            | int         | 2                    |
| Universities        | int         | 185                  |
| Sunrise             | time        | 05:28                |
| Sunset              | time        | 18:52                |
| Coastal             | bool        | TRUE                 |
| Density (pop/km²)   | float       | 6363.12              |
| UNESCO Sites        | int         | 2                    |
| Rating              | float       | 4.7                  |

Cities in the data: Tokyo, Paris, New York, Sydney, Cairo, Rio de Janeiro, Mumbai, Berlin, Cape Town, Singapore.

This template is used for:
- **Benchmarking** (`scripts/benchmark.py` reads this template, self-concatenates to 5k+ rows, and measures fill performance across backends)
- **Wide-table testing** — ensuring fillxl correctly handles tables with more than the typical 5–10 columns
- **Mixed-dtype preservation** across string, numeric, boolean, datetime, and time columns
