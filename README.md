# Final Annealing – Coils Ready Tool

Streamlit app: upload the **L2 machine report** + the **Ann list** file, and it
automatically cross-references them to find which coils finished CRM and
reached their target thickness for this stage, then gives you a styled Excel
table (same logic used in chat) with the designated Final Annealing
temperature for each coil.

## What it does
1. Reads the L2 report, takes the **last pass** of every coil.
2. Reads the Ann file, auto-detects the sheet that has `Targeted Th.` and
   `HEAT - S.T` columns.
3. Matches coils by Coil No., checks if the reached thickness (`Exit
   Thickness`) matches the stage's target thickness (`TH [mm]`) within a
   tolerance (adjustable in the app, default ±0.02mm).
4. Outputs a steel-blue styled `.xlsx` you can download directly.

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to GitHub + Streamlit Community Cloud (free, no server needed)

1. Create a new repo (or add these two files to an existing one, e.g. next to
   your `CRM-Report` repo):
   ```bash
   git init final-annealing-tool
   cd final-annealing-tool
   # copy streamlit_app.py + requirements.txt here
   git add .
   git commit -m "Final annealing ready-coils tool"
   git branch -M main
   git remote add origin https://github.com/bassamfayyad6-bit/final-annealing-tool.git
   git push -u origin main
   ```
2. Go to https://share.streamlit.io → **New app**.
3. Pick the repo, branch `main`, main file `streamlit_app.py`.
4. Deploy — you'll get a public link (e.g.
   `https://final-annealing-tool.streamlit.app`) you can open from any
   browser or phone, upload the two files, and download the result.

No secrets/API keys needed — everything runs from the uploaded files.

## Notes
- Works with `.xls` and `.xlsx` for both files.
- If your Ann file's column names change slightly, the app searches by
  partial name match (`Targeted Th`, `HEAT`, `S.T`, etc.), so small renames
  are usually fine. If it can't find a match it will show an error telling
  you what's missing.
