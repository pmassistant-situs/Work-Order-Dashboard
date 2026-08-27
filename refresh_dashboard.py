name: Refresh Work Order Dashboard

on:
  schedule:
    # Target: 9:35 AM Mountain Time daily — 5 min after the ~9:30 AM save,
    # enough buffer to avoid grabbing a half-saved file.
    # Cron doesn't auto-adjust for DST, so we schedule BOTH UTC equivalents:
    #   15:35 UTC = 9:35 AM MDT (summer, UTC-6) — correct Mar-Nov
    #   16:35 UTC = 9:35 AM MST (winter, UTC-7) — correct Nov-Mar
    # The "wrong" one for the current season just runs early/late and is
    # harmless — the correct one always lands right after your save.
    - cron: '35 15 * * 1-5'
    - cron: '35 16 * * 1-5'
  workflow_dispatch: # Also allows manual trigger from the GitHub Actions UI

jobs:
  refresh:
    runs-on: ubuntu-latest
    permissions:
      contents: write

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install pandas openpyxl

      - name: Run refresh script
        env:
          DRIVE_FILE_ID: 1xhjrKwC9bHKAdYDVfNj3TUfM09qT4RKv
        run: python refresh_dashboard.py

      - name: Commit and push updated dashboard
        run: |
          git config user.name  "Dashboard Bot"
          git config user.email "bot@thesitusgroup.com"
          git add index.html
          git diff --staged --quiet || git commit -m "Dashboard refresh $(date -u +'%Y-%m-%d')"
          git push
