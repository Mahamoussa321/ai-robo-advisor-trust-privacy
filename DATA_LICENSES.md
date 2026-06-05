# Data licenses and redistribution notes

## Study 1: Digital-twin robo-advisor experiment
The manuscript reports secondary analysis of Bonelli's public Mendeley dataset, DOI 10.17632/39rw5ywj8r.1. The full raw workbook was not available inside this project at packaging time, so Study 1 figures are reproduced from a transparent derived-effect table in `data/derived/study1/study1_effects.csv`. For a fully raw-data reproduction of Study 1, download the public workbook from Mendeley Data and add the loader to `scripts/final_robo_advisor_analysis.py`.

## Study 2: Indonesian Gen-Z robo-advisor adoption survey
Raw Excel files from the Zenodo record used in this project are included in `data/raw/study2_indonesia/`. Check the Zenodo record for the authoritative license and citation before redistributing.

## Study 3: FINRA NFCS 2024 Investor Survey
FINRA Foundation data should not be redistributed in a public GitHub repository. The private local package includes the file only so the project can be rerun on the author's computer. For public GitHub, remove the raw FINRA data and require users to download it independently from FINRA Foundation, accept the terms, and place the data file at:

`data/raw/study3_finra/NFCS 2024 Investor Data 251114.sav`

The methodology and questionnaire PDFs are included for documentation.
