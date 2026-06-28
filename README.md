# ⚜️ Scout Patrol Sorter

A free, **fully client-side** web tool that sorts Scouting America youth-training-course
participants into patrols. Everything runs in the browser — no server, no data upload —
so it's safe for youth PII and deploys to GitHub Pages with zero build step.

## What it does

Given a list of participants (**name, email, unit #, gender, birth date**), it forms patrols that follow these rules:

| Rule | Type | Behavior |
|------|------|----------|
| Patrol size **5–8** | hard | Sizes balanced as evenly as possible. Out-of-range sizes are flagged. |
| **Single-gender** patrols | hard | Male and Female split into separate patrols (M/F, Male/Female, Boy/Girl normalized). Unknown gender is listed separately. |
| **Even distribution** | soft | Patrol sizes differ by at most one. |
| **≤ 2 per unit** per patrol | soft ("if possible") | Members are swapped between similar-age patrols to separate units; remaining overages are flagged. |
| **Tent gap ≤ 730 days** | hard | Patrols are age-banded; the tool assigns **tent pairs only** (an odd leftover gets a solo tent — never a tent of three) and checks the **actual day-difference between birth dates** (default ≤730 days, applied to any tent containing a youth). |
| **Adults (18+) tent separately** | hard | Anyone **18 or older at the course date** is tented apart from youth under 18 — adults and youth never share a tent. Youth and adults are paired within their own class (a lone adult/youth gets a solo tent). |

> The tenting gap reflects the Scouting America guideline that youth sharing a tent be
> close in age (commonly applied as no more than ~2 years / 730 days apart). **Always confirm
> assignments against the current _Guide to Safe Scouting_ and Youth Protection policies**
> before publishing.

## Using it

1. Open the page.
2. **Load participants** — choose a `.csv`/`.xlsx` file, or paste rows (with a header line). Click *Load sample data* to try it.
3. **Map columns** — the tool auto-detects name/email/unit/gender/birth date; override if needed. Set the **course start date** (used to determine who is 18+ during the event) and tweak the tenting gap (days), patrol size, and per-unit options. Optionally check **Exclude adults 21 and older** to leave 21+ participants out of patrols (they're listed separately).
4. **Generate** — review patrol cards with member tables and tent assignments. Warnings highlight any rule that couldn't be fully satisfied.
5. **Adjust** — **drag participants** between patrol cards (or use the *Move…* dropdown), and **click a patrol's name to rename it**. Tent pairs and warnings recalculate automatically.
6. **Export CSV** or **Print** the results.

Input column headers are flexible (e.g. `Troop`, `DOB`, `Sex` are recognized). Dates accept
`YYYY-MM-DD`, `M/D/YYYY`, Excel dates, etc.

## Deploy to GitHub Pages

This repo is static (`index.html`, `styles.css`, `app.js`). Two options:

**A. GitHub Actions (included)** — `.github/workflows/deploy.yml` publishes on every push to `main`.
1. Create a repo (e.g. `your-name/scout-sorter`) and push these files.
2. In **Settings → Pages**, set **Source = GitHub Actions**.
3. Push to `main`; the site goes live at `https://<you>.github.io/<repo>/`.

**B. Branch source (no Actions)**
1. Push the files to `main`.
2. **Settings → Pages → Source = Deploy from a branch**, branch `main`, folder `/ (root)`.

### Push from this folder

```bash
cd D:\Scout-Sorter
git init
git add .
git commit -m "Scout Patrol Sorter"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

## Privacy

All parsing and sorting happen locally in your browser using PapaParse and SheetJS loaded
from a CDN. No participant data is ever transmitted.
