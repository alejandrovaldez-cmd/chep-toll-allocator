# CHEP Toll Allocator — Web App

Browser-based tool for processing monthly trailer-leasing toll files and producing CHEP daily bill-back allocations.

## What's in this folder

| File | Purpose |
|---|---|
| `app.py` | The Streamlit application (the actual tool) |
| `requirements.txt` | Python packages needed to run it |
| `README.md` | This file |

## Running locally (one-time setup, then daily use)

If you already have Python installed:

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`. That's it — drop your files in, click "Run Allocation," download the Excel.

## Sharing with the team (Streamlit Community Cloud)

This is the path that gets a shareable link your 6-7 finance/accounting users can bookmark.

### One-time setup

1. **Create a private GitHub repo** for this code.
   - Go to https://github.com/new → make it private
   - Name: something like `vorto-chep-toll-allocator`
   - Add your team members as collaborators

2. **Push the contents of this folder** to the repo:
   ```bash
   git init
   git add app.py requirements.txt README.md
   git commit -m "Initial commit"
   git remote add origin <your-repo-url>
   git push -u origin main
   ```

3. **Sign in to Streamlit Community Cloud** at https://streamlit.io/cloud using your GitHub account.

4. **Click "New app"** and point it at:
   - Repository: your new private repo
   - Branch: `main`
   - Main file path: `app.py`

5. **Set app visibility** in Settings → Sharing:
   - **Private** (recommended) — only people you invite by email can access. Free for 1 private app per workspace.
   - **Public** — anyone with the link. NOT recommended for this since it processes financial data.

6. **Invite your team** by email under Settings → Sharing → Viewers.

You'll get a URL like `https://vorto-chep-toll-allocator.streamlit.app/` that you can bookmark and share with the team.

### Updating the app

Make changes to `app.py`, push to GitHub, and Streamlit Cloud auto-redeploys within ~30 seconds. No re-installation for users — they just refresh their browser.

## ⚠️ Before deploying — IT/security check

This app processes sensitive financial data (vendor billing, asset records, CHEP customer info). Before deploying anywhere outside your laptop, **get sign-off from Vorto's IT/security team**. Things they'll likely want to know:

- Who can access the app (Streamlit Cloud's private mode restricts to invited Google/GitHub accounts only)
- Where data lives (Streamlit Cloud doesn't persist uploads — files exist in memory only during the session)
- What data leaves Vorto's systems (the app sends uploaded files to Streamlit's hosted environment)

If IT pushes back on external hosting, two alternatives:

1. **Run it on a Vorto-internal server** (Linux box or a Docker container behind your VPN)
2. **Each user runs it locally** on their own laptop — see "Running locally" above

## How the app works (for users)

1. **Upload the four vendor files** — Premier, Star, XTRA, Bestpass (Excel)
2. **Upload the Master Asset Document** (only needed for MAD-driven mode)
3. **In the sidebar:**
   - Pick processing mode (MAD-driven for full automation, Pre-mapped to use accounting's existing columns)
   - Set the cost month (vendor invoice month)
   - Set the billing month (the month being billed to CHEP)
   - Toggle "Filter by Invoice Date" (default ON — matches accounting's methodology)
4. **Click "Run Allocation"**
5. **Review the results inline:**
   - Per-vendor breakdown
   - Allocation by CHEP market
   - Exceptions (if any)
   - Deferred rows (if any)
6. **Click "Download Excel"** to get the daily allocation file

## Troubleshooting

**"Failed to load Premier" (or any vendor)**
- Check that the file structure matches the expected format (e.g., XTRA's headers are on row 3, Star's are on row 1)
- Vendor file templates can change month-to-month — if a column rename happens, the loader functions in `app.py` need an update

**"No CHEP toll data found"**
- Check the cost month/year settings match your vendor files' invoice dates
- Try toggling off "Filter by Invoice Date" to see if that's excluding everything

**"Units not in MAD" warning**
- Bestpass sometimes records driver-incurred tolls under generic Unit IDs like "25" or "28". These are real and should be excluded from CHEP billing.
- If a real trailer ID isn't matching, it might mean the MAD export is stale — re-export from Google Sheets

## Methodology notes

- The app uses the same logic as the standalone Python scripts (`chep_toll_allocator_perday.py`, `chep_toll_allocator_mad.py`) documented in the CHEP Toll Allocation Process Guide
- Region → CHEP Market mapping uses three-tier fallback: Current Region → Originating Market → Bestpass Cost Center / SCAC code
- The 33 CHEP Dedicated Lane markets are encoded in the app
- Per-row overrides are applied for the known Virginia mislabel issues
