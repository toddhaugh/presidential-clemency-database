# Build report

- Workbook: `data/ICGE Presidential Clemency Database - Nov 13 2026VF.xlsx`
- Built: 2026-11-20 19:36 UTC
- Records in workbook: 4,960; named recipients on Search records: 3,401
- Data through: Oct. 13, 2026; last checked against DOJ: Nov. 13, 2026

## Cross-check against the workbook's Website Map sheet

All 389 figures the Website Map lists match the workbook's own formulas.

## Statements in the page text

These sentences are written into the pages. If one is marked FALSE, the wording needs updating in `src/pages`.

- OK – Classification page: drug trafficking and fraud/theft/embezzlement "together account for more than half" of individual grants
- OK – Classification page: drug offenses are the largest group of non-white-collar grants
- OK – Classification page: fraud/theft/embezzlement is "the most common" white-collar offense type
- OK – Classification page: the only category combinations not listed in Table 5 involve unethical business practices
- OK – By administration page: "Most grants came in the final year" (Clinton second term)
- OK – By administration page: "Most grants came in the final year" (Trump first term)

## Text that is not computed

Update these by hand in `src/pages/*.template.html` when they change: the Methodology change log and "Last revised" date, the count of 17 records too vague to classify, the "about 1,500" unnamed Jan. 6 recipients, the 77 Nov. 7, 2025 proclamation pardons, the cleaning history, and the news item on the About page.
