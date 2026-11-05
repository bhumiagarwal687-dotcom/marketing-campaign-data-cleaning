# marketing-campaign-data-cleaning
Excel data cleaning project on a Kaggle marketing campaign dataset: removed duplicates and outliers, fixed dates and categories, filled missing income, and added analysis-ready columns.
# Marketing Campaign Data Cleaning (Excel)

Cleaning of a customer marketing campaign dataset in Excel, preparing it for analysis and dashboards.

## Repository contents

| File | Description |
|---|---|
| `marketing_campaign_corrected.xlsx` | Workbook with two sheets: `Clean_Data` (final cleaned data) and `Raw_Data` (original, untouched) |
| `README.md` | This file |

## Dataset overview

- **Source:** downloaded from Kaggle (add the link to the dataset page here and credit the original author)
- **Original:** 2,240 customers, 29 columns
- **Cleaned:** 2,053 customers, 33 columns
- **Covers:** customer demographics (birth year, education, marital status, income, children), spending by product category (wine, fruit, meat, fish, sweets, gold), purchases by channel (web, catalog, store), and responses to marketing campaigns
- **Customer join dates:** July 2012 to June 2014

## Data issues found and how they were fixed

| # | Issue | Action | Rows affected |
|---|---|---|---|
| 1 | Duplicate customers: rows identical in every column except `ID` | Removed (kept first occurrence) | 182 |
| 2 | Impossible income of 666,666, far above all other values | Removed | 1 |
| 3 | Impossible birth years (1893, 1899, 1900) | Removed | 3 |
| 4 | Inconsistent categories in `Education` | `2n Cycle` merged into `Master` | 203 (before duplicate removal) |
| 5 | Inconsistent / invalid categories in `Marital_Status` | `Alone` changed to `Single`; `Absurd` and `YOLO` changed to `Other` | 3 and 4 |
| 6 | Duplicates that only appeared after standardising categories | Removed (one customer matched another once `Alone` became `Single`) | 1 |
| 7 | `Dt_Customer` stored as text (dd-mm-yyyy) | Converted to real Excel dates | all rows |
| 8 | Missing `Income` | Filled with the median income of the same education level; flagged in `Income_Imputed` | 24 |
| 9 | `Z_CostContact` and `Z_Revenue` hold the same value in every row | Dropped (no information) | 2 columns |

**Order matters:** categories are standardised *before* the final duplicate check, otherwise rows that are really the same customer are missed.

## New columns added

| Column | Definition |
|---|---|
| `Income_Imputed` | `Yes` if the income was filled in, otherwise `No` |
| `Age` | 2014 minus `Year_Birth` (the data ends in mid-2014) |
| `Total_Children` | `Kidhome` + `Teenhome` |
| `Total_Spend` | Sum of the six `Mnt...` spending columns |
| `Total_Purchases` | Web + catalog + store purchases |
| `Campaigns_Accepted` | Sum of `AcceptedCmp1` to `AcceptedCmp5` plus `Response` |

## How to reproduce in Excel

1. Copy the raw sheet and work on the copy.
2. **Data > Remove Duplicates**, leaving `ID` unticked.
3. Convert `Dt_Customer` to dates: `=DATE(RIGHT(H2,4),MID(H2,4,2),LEFT(H2,2))`, then paste as values.
4. Use **Find & Replace** to standardise `Education` and `Marital_Status`.
5. Sort `Income` and `Year_Birth` to find and delete impossible values.
6. Add `Income_Imputed` (`=IF(E2="","Yes","No")`), then fill blank incomes with the median for the same education level.
7. Delete `Z_CostContact` and `Z_Revenue`.
8. Add the calculated columns above and check for blanks and duplicates.
9. Run a final duplicate check (excluding `ID`) after all standardisation.

## Final quality checks

- No blank cells
- No duplicate IDs or duplicate rows
- Income range: 1,730 to 162,397
- Age range: 18 to 74
- All dates between July 2012 and June 2014

## Tools

Microsoft Excel (Remove Duplicates, Find & Replace, sorting and filtering, formulas).

## Assumptions and limitations

- Missing incomes were estimated with a group median, which is an approximation; use the `Income_Imputed` flag to exclude them if needed.
- Removing outliers and duplicates assumes they are data-entry errors.
- `Age` is calculated relative to 2014, not today's date.
