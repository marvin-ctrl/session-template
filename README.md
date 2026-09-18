# Futsal Session Coach

## Session plan template (PowerPoint)

`futsal_session_template.pptx` is a seven-slide, fully editable session-plan
system for a national-team futsal programme. Every element is a native
PowerPoint shape — no images, nothing locked. `futsal_session_template.pdf`
is the same file rendered for quick viewing.

| # | Slide | Use |
|---|-------|-----|
| 1 | Master template | Keep clean; never edit directly |
| 2 | Working copy | Duplicate this every week and fill it in |
| 3 | Annotated guide | Same layout with deletable blue "HOW TO" notes |
| 4 | Component library | Court-scale players, body-shape pucks, vision cones, scan zones, arrows, zones, grouped courts — copy onto any diagram |
| 5 | Session control board | Time-proportional load timeline with a draggable intensity curve, game-moment fingerprint, rotation board, review loop, audit, kit list |
| 6 | Set pieces | Futsal's own restarts and phases: kick-in, corner, goal clearance, powerplay 4v3, defending powerplay, 10 m accumulated foul |
| 7 | On-court cards | Print, cut along the dashed lines, pocket — one cue card per block |

**Recolour the whole system in one move.** The three identity colours live in
the presentation theme: Design → Variants → Colours → Customise, then set
Accent 1 (your team), Accent 2 (opposition) and Accent 3 (court lines).
Every slide updates at once. Body text stays black on purpose.

**Regenerate from source**

```bash
pip install -r requirements.txt
python3 generate_template.py
```

The generator is `generate_template.py`; it renumbers shape ids, wires colours
to the theme and fails the build if any duplicate id survives.

---

## Courtside app

Mobile-first courtside session manager for Kauri Futsal-style pickup games.

The app runs as a Vite React PWA, stores data locally, generates fair teams, tracks rests, records results, and keeps session history. There is no backend, authentication, or external database in this version.

## Commands

```bash
npm install
npm run dev
npm test
npm run build
```

## Data

Local data is stored under the browser key `futsal-coach-session-v1`. Use the Backup tab to export or import JSON.
