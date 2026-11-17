# NCL 60 & 90 Days Cases — web reports

Web version of `GENERATE_REPORT_60_AND_90_DAYS_CASES.xlsm`. It is a static site
(HTML + JavaScript only), so it runs on GitHub Pages with no server.

## Reports

For both 60 DAYS and 90 DAYS cases:

| Web report | Sheet in the Excel workbook |
|---|---|
| Abstract (SD wise & PS wise) | `ABSTRACT_60` / `ABSTRACT_90` |
| UI < 45 / < 75 days | `UI<45_DAYS` / `UI<75_DAYS` |
| UI 45–60 / 75–90 days | `UI_BET_45_TO_60_DAYS` / `UI_BET_75_TO_90_DAYS` |
| UI > 60 / > 90 days | `UI_MORETHAN_60_DAYS` / `UI_MORETHAN_90_DAYS` |
| All cases | `NCL_60` / `NCL_90` |

- Days are counted up to today each time the page is opened (change **As on** to see another date).
- Click any number in the abstract to see the cases behind it.
- **Download Excel** saves the report on screen; **Download all reports** saves one workbook with all ten sheets;
  **Print / PDF** prints on Legal paper (choose "Save as PDF" in the print window).

## Put it on GitHub Pages (one time)

1. Create a repository on GitHub (for example `chittoorncl`) and upload **everything in this folder**
   (`index.html`, `assets`, `vendor`, `data`, `.nojekyll`).
2. Repository → **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main`, folder `/ (root)` → Save.
3. After a minute the site opens at `https://<user>.github.io/<repository>/`.

## Admin login and upload

GitHub Pages cannot store passwords or files itself, so the admin signs in with a GitHub
**access token** for this repository. Uploading saves the data into the repository
(`data/ncl60.json`, `data/ncl90.json`) and GitHub republishes the site.

Create the token once:

1. GitHub → your photo → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
2. Repository access: **Only select repositories** → choose this repository.
3. Permissions → Repository permissions → **Contents: Read and write**.
4. Generate, copy the token (`github_pat_…`) and keep it safe. Anyone with this token can change the data.

Updating the reports:

1. Open the site → **Admin** → enter the token → **Sign in**.
2. Choose the `60 DAYS` and/or `90 DAYS` Excel file (`.xlsx` or `.xls`, columns A:H as exported).
   The reports behind the window show a preview.
3. Press **Publish**. Everyone sees the new data after about a minute.

## Settings (`assets/config.js`)

- `subDivisions` — police stations under each sub-division (the SD sheet). Add or move stations here.
- `excludeStations` — stations dropped while importing (Punganur UPS, Sodam, Somala, Chowdepalli).
- `fromDate`, `unitName`, `appTitle`, `credit` — headings.
- `github` — only needed if the site uses a custom domain or is not at the top of the repository.

## Notes

- If the repository is public, the case data in `data/` can be read by anyone who has the link.
- Rules are the same as the macros: a case with a chargesheet date is FR/CS within or after 60 (90) days
  of the FIR date; Court Disposal without a chargesheet date is `LAD_AT_UI_STAGE`; other cases without a
  chargesheet date are UI, grouped by days from the FIR date to today.
- To test on your own computer, run `python -m http.server` in this folder and open `http://localhost:8000`
  (opening `index.html` by double-click will not load the data).
