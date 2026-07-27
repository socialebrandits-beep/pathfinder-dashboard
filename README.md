# Pathfinder Lead Dashboard — daily auto-refresh

This makes your dashboard update itself every day, for free, with zero manual work and zero Claude usage after setup.

## How it works

- `template.html` — the dashboard page (same look as before), with a placeholder where the numbers go.
- `build_dashboard.py` — downloads both live Google Sheets, recomputes the exact-match metrics (contacted / followed up / not contacted, same rules we finalized), and writes the numbers into `index.html`.
- `.github/workflows/update.yml` — tells GitHub to run that script automatically once a day and publish the result. Runs on GitHub's free infrastructure — not this Claude session, so it costs nothing and needs nothing from you.

## One-time setup (5 minutes)

1. In your existing GitHub repo (the one you already made for `index.html`), add these 3 items exactly as named, keeping the folder structure:
   - `template.html` (repo root)
   - `build_dashboard.py` (repo root)
   - `.github/workflows/update.yml` (this exact path — the `.github/workflows/` folder matters, GitHub looks for it there)
2. Delete the old plain `index.html` you uploaded manually — the workflow will generate a fresh one the first time it runs.
3. Go to your repo's **Settings → Actions → General**, scroll to "Workflow permissions," and select **"Read and write permissions."** (This lets the daily job commit the updated file back to your repo. One-time toggle.)
4. Go to the **Actions** tab of your repo → click "Update Pathfinder Dashboard" → click **"Run workflow"** to trigger the first run manually and confirm it works. After ~30 seconds you should see a green checkmark and a new commit updating `index.html`.
5. Your GitHub Pages URL (Settings → Pages) will now show the freshly-generated dashboard, and will keep refreshing automatically every day at 00:00 UTC — no further action needed, ever.

## Changing the daily time

Open `.github/workflows/update.yml` and edit this line:

```
- cron: '0 0 * * *'
```

The two numbers are `minute hour` in UTC. For example, `'30 1 * * *'` runs at 01:30 UTC. You can also just click "Run workflow" anytime in the Actions tab for an on-demand refresh.

## If a run ever fails

Check the Actions tab — a red X means something changed on the source sheets (e.g. a sheet got renamed, or a column got moved). The most likely cause is a sheet or tab name changing; `build_dashboard.py` looks up sheets by their exact tab name, so a rename there is the one thing that would need a matching one-line update in the script.

## Privacy note

The generated dashboard still includes lead names and phone numbers in the "Recent activity by date" section, and phone numbers link out to the source Google Sheet. Since your GitHub Pages site is public, anyone with the URL can see that data. If you'd rather not expose it, let me know and I'll adjust the script to drop names/phone numbers from the public version (aggregate counts only) — just say the word.
