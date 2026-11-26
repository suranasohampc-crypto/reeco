# GST REECO

Upload these files to a private GitHub repository. In Streamlit Community Cloud select the repository, branch main, and app.py. Verify private access before uploading audit data. Do not commit client workbooks.

## Password protection

In the deployed app's Streamlit settings, open Secrets and add `APP_PASSWORD = "your unique password of at least 16 characters"` (replace the example). Save and reload the app. Never commit the real password to GitHub. The app stays locked until configured. Keep the Streamlit app private as well; this is a shared password, not individual user authentication.

## Enable or disable the password screen

In Streamlit Settings → Secrets, set `PASSWORD_PROTECTION = true` to enable the password screen or `PASSWORD_PROTECTION = false` to disable it. Use unquoted TOML booleans. Keep `APP_PASSWORD` configured so you can re-enable protection later. If the toggle is omitted, protection is on. This does not change Streamlit's private app setting.

## Batch uploads

Select multiple offset summaries. Download each workbook or all successful workbooks as a ZIP. Invalid inputs are reported separately.

## Verification preview

Deploy `packages.txt` alongside the app on Streamlit Community Cloud. For local or other hosting, install LibreOffice Calc and make `libreoffice` or `soffice` available on PATH. Recalculation uses a separate temporary directory and LibreOffice profile per workbook; a workbook is not offered for download if recalculation or numerical verification fails.

The preview checks the generated workbook against the uploaded source by month and tax head. It reads annual totals from the exact recalculated workbook offered for download and compares them to sums of mapped source months, not to the source's printed total row. Blank input and numeric zero are distinct. A missing GSTIN or duplicate GSTIN/year is flagged for manual review. These are import checks only; books figures and client identity against an external register are not verified. No email is sent.
