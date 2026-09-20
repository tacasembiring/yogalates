# Yogalates

A single-page, animated yoga landing page. No build step, no dependencies — everything lives in [index.html](index.html).

## What's inside
- Animated gradient mesh background (four drifting color blobs)
- A **breathing orb** at the center with a meditating figure — expands/contracts on an 8s inhale/exhale rhythm
- Floating lotus petals, a pulsing "heart" in the headline, and a glowing gradient headline
- A center call-to-action button with a sheen sweep, pulsing halo rings, and a click ripple
- Fully responsive (phone → laptop), safe-area aware, and respects `prefers-reduced-motion`

## Preview locally
Just open the file — or serve it:
```bash
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Deploy to GitHub Pages
1. Push this repo to GitHub.
2. Repo **Settings → Pages**.
3. Under **Build and deployment**, set **Source: Deploy from a branch**.
4. Choose your branch (e.g. `main`) and folder `/ (root)`, then **Save**.
5. Your page goes live at `https://<username>.github.io/<repo>/` within a minute or two.

The included `.nojekyll` file tells Pages to serve the site as-is.
