# Ghostladder Dashboard — Project Context

## What This Is
Ghostladder is a real-time price guide and ROI calculator for ghostwrite blind box toys (MLB, NBA, etc.). It tracks prices across blind box sets, builds statistical price estimates where sales data is thin, and calculates ROI. Think Campless but for ghostwrite.

The entire project is a **static HTML/CSS/JS dashboard** — no build system, no framework, no bundler. Each page is a self-contained `.html` file. Shared styles live in `nav.css` and `variants.css`.

---

## Pages

| File | Route | Description |
|------|-------|-------------|
| `index.html` | `/` | Sales Intel — main market dashboard |
| `guide.html` | `/guide` | Price Guide — per-set price tables |
| `listings.html` | `/listings` | Listings Intel — live eBay listings |
| `sets.html` | `/sets` | Sets — ROI calculations per set |
| `heatmap.html` | `/heatmap` | Set Map — estimation model explorer |
| `case-sim.html` | `/case-sim` | ROI Simulator — open a virtual case |
| `explainer.html` | `/explainer` | How Ghostladder works (editorial) |
| `gl50.html` | `/gl50` | GL50 — top 50 players ranking (BETA) |
| `player-index.html` | `/player-index` | Player Indexes — sub-page under GL50 |
| `sales-feed.html` | `/sales-feed` | Raw sales feed |
| `model.html` | `/model` | Model internals |
| `report.html` | `/report` | Report page |
| `post-builder.html` | `/post-builder` | Post builder tool |
| `admin.html` | `/admin` | Admin panel |
| `admin_slides.html` | `/admin_slides` | Admin — slides |
| `admin_views.html` | `/admin_views` | Admin — analytics |

---

## Nav Structure (applies to ALL pages)

```html
<a href="index.html">Sales Intel</a>
<a href="guide.html">Price Guide</a>
<a href="listings.html">Listings Intel</a>
<a href="sets.html">Sets</a>
<a href="heatmap.html">Set Map</a>
<a href="case-sim.html">ROI SIM</a>
<div class="nav-dropdown">
  <a href="gl50.html">GL50 <span class="beta-badge">beta</span></a>
  <div class="nav-dropdown-menu">
    <a href="player-index.html">Player Indexes</a>
  </div>
</div>
```

- `class="active"` goes on whichever link matches the current page
- For `gl50.html`: active on the `<a href="gl50.html">` inside the dropdown wrapper
- For `player-index.html`: active on the `<a href="player-index.html">` inside the dropdown menu
- Nav dropdown CSS lives in `nav.css` (`.nav-dropdown`, `.nav-dropdown-menu`)
- **Nav changes always go in `nav.css`** (shared by all pages) and the HTML updated across all pages

---

## CSS Design System

All pages use the same CSS custom properties defined in `:root`:

```css
--bg: #080808;
--s1: #101010;      /* surface 1 */
--s2: #161616;      /* surface 2 */
--s3: #1e1e1e;      /* surface 3 */
--border: #282828;
--accent: #e8ff47;  /* yellow-green */
--accent-dim: rgba(232,255,71,0.12);
--red: #ff4040;
--green: #3ddc84;
--text: #f0f0f0;
--muted: #666;
--muted2: #999;
--mono: 'DM Mono', monospace;
--display: 'Bebas Neue', sans-serif;
--serif: 'Instrument Serif', serif;
```

Fonts loaded from Google Fonts: Bebas Neue (display/headings), DM Mono (body/data), Instrument Serif (editorial prose).

---

## Backend: Supabase (PostgREST)

- **Project URL:** `https://sabzbyuqrondoayhwoth.supabase.co`
- **Anon key:** `sb_publishable_KAbUO5YvOapq_kWkHqt7BA_UTwrP60s`
- All DB calls use raw `fetch()` against the PostgREST REST API — no Supabase JS client

Standard helpers (defined per-page, not in a shared file):
```js
const SB_URL = 'https://sabzbyuqrondoayhwoth.supabase.co';
const SB_KEY = 'sb_publishable_KAbUO5YvOapq_kWkHqt7BA_UTwrP60s';
const SB_H   = { 'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY, 'Content-Type': 'application/json' };

async function sbGet(table, qs)         // GET /rest/v1/{table}?{qs}
async function sbPatch_(table, qs, body) // PATCH
async function sbPost_(table, body)      // POST
```

**PostgREST gotcha:** Item IDs like `v1|358372033719|0` contain pipe chars (`|`) that must be `encodeURIComponent`-encoded in URLs. PostgREST decodes `%7C → |` server-side when matching. Always encode when building `in.(...)` filters.

---

## Google Apps Script (`apps-script-functions.js`)

Runs as a GAS project connected to a Google Sheet. Key functions:

- `refreshListingsScrapedAt_(itemIds, runId)` — updates `scraped_at` on listings rows. Chunk size = 50 (keeps URL under GAS's ~2048 char limit; encoded pipe IDs are ~22 chars each).
- `classifyLiveListings_` — Haiku-based batch classifier (new version); old JS version renamed `classifyLiveListings_JS_` with deprecation comment (do not delete).

---

## Figure Images

Cards use player figure PNGs at:
```
assets/figures/{set-slug}/{player-slug}_{variant-slug}.png
```

Set slug: lowercase, spaces → hyphens, strip special chars (e.g. `MLB 2025 Game Face` → `mlb-2025-game-face`)
Player slug: lowercase, strip accents/punctuation, spaces → underscores (e.g. `Ronald Acuña Jr.` → `ronald_acuna_jr`)
Variant slugs: `base`, `gold`, `chrome`, `emerald`, `fire`, `victory`, `vides`, `sp`, `bet-on-women`

All `<img class="box-card-figure">` tags include `onerror="this.style.display='none'"` — missing images silently hide.

---

## case-sim.html — Simulator Rules

When opening a virtual case:
1. **Step 1:** Pick one fully-random hit (non-base card) from the hit pool
2. **Step 2:** Pick `casePack - 1` base cards where **every card must be a unique player** — no player appears twice in the grid. Uses `pickUniqueBasePulls(remainingPool, hitPlayer, count)`.

The hit ratio and PWAHV calculations are unaffected by this display rule.

---

## explainer.html — Layout Hierarchy

```
.post-title     — Bebas 58px  — H1 (page title)
.section-title  — 30px        — H2 (major sections)
.section-sub-title — 26px     — H3-level sub-sections
.section-eyebrow — 8px caps   — label above a section
.prose          — body text
.post-lead      — larger intro paragraph
```

Custom classes added for explainer:
- `.section-sub-title` — 26px Bebas, used for "How We Fill the Gaps" and "Three Modes, One Toggle"
- `.cta-row` — 3-column grid of `.cta-nav-btn` cards at the bottom of the page

---

## Key Conventions

- **Stat display:** Always use `.toFixed(2)` for hit ratio and bases-per-case stats (e.g. `(+stats.hitRatio).toFixed(2)`)
- **No build step:** Edit `.html` files directly — changes are live immediately
- **Shared CSS only in `nav.css` / `variants.css`** — page-specific styles go in `<style>` blocks inside each page's `<head>`
- **Admin subnav** (`admin.html`, `admin_slides.html`, `admin_views.html`) — secondary nav bar below the main header; does NOT appear on public-facing pages (including `listings.html`)
- **BETA badge** — `<span class="beta-badge">beta</span>` — currently on GL50 only; not on ROI SIM

---

## Pending / In-Progress Work

- **Global nav update** — nav.css is done. Still need to apply updated nav HTML to all 16 pages (some may already be done — grep for `Listings Intel` to check which pages are updated)
- **GAS:** `classifyLiveListings_JS_` deprecation comment + new Haiku-based `classifyLiveListings_` still needed
- **player-index.html enhancements (deferred):**
  - Move "Performance Comparison" chart to bottom
  - Add "Total Sales" stat to detail panel
  - Date range toggle should update all detail tabs
  - Swap Caitlin Clark / Ohtani order — Ohtani first
