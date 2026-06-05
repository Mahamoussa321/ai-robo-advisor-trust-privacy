# Personalization, Privacy, and Trust in AI Robo-Advisory

Reproducible project package for the three-study manuscript:

**Personalization, Privacy, and Trust in AI Robo-Advisory: Experimental Evidence, Real-World Survey Validation, and Machine-Learning Prediction**

## What this project reproduces

Running the code regenerates:

- all analysis tables in `outputs/*.csv`
- all colorful manuscript figures in `outputs/*.png`
- machine-learning performance and feature-importance tables
- the final Word manuscript in `manuscript/robo_advisor_final_submission_ready.docx`

## Folder structure

```text
data/
  raw/study2_indonesia/        # Zenodo robo-advisor adoption survey Excel files
  raw/study3_finra/            # FINRA data for private local use only
  derived/study1/              # transparent Study 1 effect table
scripts/
  final_robo_advisor_analysis.py
  create_final_submission_docx.py
  run_all.py
  verify_reproducibility.py
outputs/                       # regenerated tables and figures
manuscript/                     # regenerated manuscript
logs/                           # run logs
```

## Quick reproducibility check

From the project root:

```bash
python -m venv .venv
.venv\Scripts\activate     # Windows PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python scripts/run_all.py
```

On Mac/Linux, activate with:

```bash
source .venv/bin/activate
```

## Important GitHub warning

Use the **PUBLIC GitHub package** for pushing to GitHub. Do not push the private package if it contains raw FINRA files. FINRA data must be downloaded independently by each researcher under FINRA Foundation terms.

## GitHub commands

```bash
git init
git add README.md DATA_LICENSES.md LICENSE_CODE.txt requirements.txt environment.yml .gitignore scripts data/raw/study2_indonesia data/derived/study1 outputs manuscript
git commit -m "Initial reproducible robo-advisor analysis package"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

Do not add `data/raw/study3_finra/NFCS*Data*` to a public repository.

## Suggested repository name

`ai-robo-advisor-trust-privacy-ml`

## Citation notes

Cite the original datasets in the manuscript and README before submission. The code is reproducible, but the final journal upload still requires author affiliations, corresponding-author information, funding statement, conflict-of-interest statement, and target-journal formatting.
