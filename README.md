# Funnel Cost Calculator, K&K Media

Lead magnet for LinkedIn post Numbers first #6 (comment FUNNEL). Shows what each step from click to order costs in ad dollars, and what +1 percentage point on each step is worth.

- Static site: Vite + React + TypeScript, deployed to GitHub Pages by `.github/workflows/deploy.yml`.
- No tracking, no cookies, no network requests (strict CSP). Inputs stay in localStorage.
- `npm test` checks the math against the post: 10,000 clicks at $1.50, 85% / 8% / 50% / 60% gives 8,500 / 680 / 340 / 204, costs $1.76 / $22.06 / $44.12 / $73.53, and 8.8% view to cart gives 224 orders at a $66.96 CPA.
- Download Excel builds a branded workbook with live formulas (exceljs, loaded on demand).
- `npm run check:dashes` fails the build on en or em dashes in interface text.
