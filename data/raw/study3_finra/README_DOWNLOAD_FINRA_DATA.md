# Download FINRA NFCS data here

This public GitHub-ready package intentionally does not include the raw FINRA NFCS Investor Survey data.

To reproduce Study 3:

1. Go to the FINRA Foundation NFCS Data and Downloads page.
2. Download: 2024 National Financial Capability Study -> Investor Survey -> Data & Data Info.
3. Accept FINRA's terms.
4. Place the CSV-format data file here with this exact filename:

   `NFCS 2024 Investor Data 251114.sav`

In the uploaded file used for this project, the extension was `.sav` but the content was UTF-8 CSV text; the analysis script reads it with `csv.DictReader`.
