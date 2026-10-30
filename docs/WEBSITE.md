# Project website

The static K-SFTs landing page lives in `docs/index.html`, with its styles in
`docs/assets/site.css`. It reuses the existing project illustration and CCR
framework diagram. Its content is authored for this repository.

## Publishing

The page is published at <https://danielep71.github.io/K-SFTs/>, by GitHub
Pages set to **Deploy from a branch**, branch `main`, folder `/docs`; each
push to `main` redeploys it. `docs/.nojekyll` makes this a plain static site. No package
installation, JavaScript, build step, analytics or third-party fonts are needed.
`docs/assets/SFT_social_preview.html`, the source of the hero image, is
published with the folder but not linked; only it loads Google Fonts, when
opened ([tools/README.md](../tools/README.md#hero-image)).

Changes follow the repository task/release/main PR process. The live site
updates only when the page changes reach `main` and the Pages deployment
succeeds. Adding a README link alone does not establish a live deployment.

## Maintenance

Keep prototype claims aligned with `methodology/PROTOTYPE.md` and the
methodology of record. Keep evidence links scoped to their recorded revision;
do not turn a static check or an illustrative image into a validation claim.
Update the canonical and social-preview URLs if the hosting location changes.

Before publishing, run the repository checks, verify local image paths and
fragment links, inspect the responsive layout, and confirm the live page after
GitHub Pages finishes its build. The website and wiki links appear together,
in that order, in the root README.
