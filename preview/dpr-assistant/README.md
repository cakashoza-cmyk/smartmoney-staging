# DPR Assistant (preview / staging)

Plain questions in, bank-ready **DPR + CMA Excel** out. Static site; runs in the browser.

- **Engine:** the owner's 17-tab DPR/CMA model (Rs lakh). `template_public.xlsx` is a sanitised copy: all original borrower inputs removed, hard-coded names replaced by formulas, and the formula fixes below applied. Users' answers are written only into the blue input cells (map: `input_map.json`, 563 blue cells) by patching the sheet XML in the browser (`xlsxfill.js`), so every formula, style, chart and dropdown is kept.
- **Extra sheets added by the app (no existing formula reads them):** `18_Profile_&_Collateral` (promoters, activity, project cost by type, means of finance, collateral), `19_Unknown_Collect` (every unanswered input + soft warnings + defaults used), `20_Report_Narrative` (optional).
- **Number rules:** figures come only from the user's answers, from documented defaults shown in the app (tax rate, margins, BEP split, stress shocks, depreciation rates), or from simple arithmetic on the user's own answers shown back to them. Blank answers stay blank and are listed. Warnings never change a value.
- **Privacy:** answers are kept in the browser's localStorage. Making the Excel sends nothing anywhere. Optional AI help is opt-in per use.
- **AI (optional):** `worker/` is a Cloudflare Worker that holds the DeepSeek key as a secret. It has only two fixed prompts (`/extract`: words to fields, which the user confirms; `/narrative`: report text with `{{PLACEHOLDERS}}` instead of figures). It checks the origin, rate-limits per IP and per day, limits request size and does not log or store request bodies. The app redacts names, PAN, Aadhaar, account numbers, phones and emails before sending, and drops any AI sentence that contains a number not found in the user's own words. The placeholders become live Excel formulas, so figures in the report always come from the workbook. If the proxy is down, the deterministic flow works unchanged.

## Template fixes applied in `template_public.xlsx`
| Where | Problem | Cells changed |
|---|---|---|
| 12_Balance_Sheet E14:N14, E17:N17 | IF(k==3,...) invalid syntax (Err:510) broke projected USL/capital, TNW, cash and BS check | 20 |
| 1_DPR_Cover_Note A54,F54; 2_Executive_Summary A1 | hard-coded names of the original (real) borrower and its auditor | linked to Tab 3 C5/C7 |
| 4_Historical_&_Benchmarks G38:I42, G52:I52 | % of total assets divided by row 53 (non-current loans) not row 54 (total assets) | 18 |
| 10_Unit_Economics C10:L10 | 90% yield multiplied volume by 0.90 instead of dividing (display row only) | 10 |
| 8_Debt_Schedules fresh-disbursement rows (D:M) | existing TL entered with start date = 1 Apr of projection year 1 was counted twice (ghost debt on BS) | 100 |
| 7_Depreciation_&_FACR row 38 | FACR exploded on floating-point TL residue after last instalment | 11 |
| 17_Sensitivity_Analysis L10:L14 | AVERAGEIF(>0) silently dropped failing (zero/negative) stress years | 5 cells |

## Known limitations (not changed)
- `6_Capex_&_Means` F32 adds existing unsecured loans (Tab 4 D40) to the project's means of finance, so the promoter "plug" is understated (Apex: USL 35.0 / plug 54.6 vs answer key 20.0 / 69.6).
- Tab 6 F34 "Subsidy / Capital Grant" is a fixed 0. Subsidy is recorded on the profile sheet only.
- Project cost covers fixed assets only; margin money for WC is a memo on the profile sheet.
- No dividends/drawings in projections; cash builds up while the CC stays fully drawn.
- Average DSCR on Tab 15 (weighted) differs from the simple average used on Tab 17.
- Only Tandon Method 1/2 MPBF; turnover (Nayak) method not built in.
- GST on capex is capitalised.

Fictional demo case: "Apex Precision Components" (practice case; not a real borrower).
