# Presidential Clemency Database

The public website of the ICGE Presidential Clemency Database: every presidential pardon and commutation since January 20, 1989, classified as white-collar or not under one published rubric.

The site has no server and no database. It is six web pages and one data file (`records.json`), published from this repository by GitHub Pages. Every figure on the pages is calculated from the ICGE workbook, so the site and the workbook cannot drift apart.

## Updating the data (the routine job)

1. Update the workbook in Excel as usual. The **Clemency Data** sheet is the only data the site reads. It also reads the date next to "Last checked against DOJ" on the **README** sheet, so update that date when you check DOJ.
2. In this repository on github.com, open the **data** folder, click **Add file → Upload files**, drag in the workbook, and click **Commit changes**. You don't need to rename it or delete the old one: the most recently uploaded workbook is always the one used.
3. Wait two or three minutes. The **Actions** tab shows the rebuild:
   - **Green check:** the site is updated. Open `BUILD_REPORT.md` to see what was built and the checks it ran.
   - **Red X:** nothing on the site changed. Click the run to see why. The message names the sheet, row and column to fix, for example an unknown Data Status or a missing date. Fix the workbook and upload it again.

GitHub Pages then publishes the new version, usually within a minute or two more.

## What the build checks

- Every row of Clemency Data has a valid record ID, date, term, clemency type, status, group, white-collar answer, confidence and categories. Any problem stops the build, so bad data never goes live.
- Every figure the workbook's **Website Map** sheet lists is compared with the figure the site will show. Any difference is listed at the top of `BUILD_REPORT.md`.
- Sentences in the page text that depend on the data (for example "Most grants came in the final year") are re-tested. Any that are no longer true are marked **FALSE** in the report.

## Changing wording or design

The page designs are in `src/pages/` (`index` is the Overview page). Each page has two files:

- `<page>.template.html`: the page itself. Edit the wording here. Text like `{%wc_ind%}` or `{%json:years%}` marks a figure the build fills in, so leave those as they are.
- `<page>.shell.html`: the packaged fonts, logo and page engine. Don't edit it.

Committing a change to a template rebuilds the site automatically. A few things are typed in by hand and need updating when they change. They're listed at the bottom of `BUILD_REPORT.md`: the Methodology change log and "Last revised" date, the count of 17 records too vague to classify, the "about 1,500" unnamed Jan. 6 recipients, the 77 Nov. 7, 2025 proclamation pardons, the cleaning history, and the news item on the About page.

## Files

| Path | What it is |
| --- | --- |
| `data/*.xlsx` | The ICGE workbook(s). The most recently uploaded one is used. |
| `index.html` … `about.html`, `records.json` | The published site. **Generated, so don't edit these directly.** |
| `BUILD_REPORT.md` | Results of the last build. |
| `src/pages/` | Page designs (templates). |
| `scripts/build.py` | Turns the workbook into the site. |
| `.github/workflows/build.yml` | Runs the build automatically on GitHub. |

To build on your own computer instead: `pip install -r scripts/requirements.txt`, then `python scripts/build.py`.
