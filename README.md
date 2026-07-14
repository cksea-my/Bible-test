# Excel-Driven Quiz for GitHub Pages

This is a static HTML/CSS/JavaScript quiz. Quiz content is loaded from `data/quiz_questions.xlsx`.

## Update questions
1. Open `data/quiz_questions.xlsx`.
2. Edit questions or add rows using the existing headers.
3. Use one worksheet per category. Duplicate an existing category sheet when adding a category.
4. Enter `A`, `B`, `C`, or `D` in **Correct Answer**.
5. Enter `Yes` in **Active**, save, and replace the workbook in GitHub.

Do not rename the workbook unless you also update `CONFIG.workbook` in `js/app.js`.

## Publish with GitHub Pages
1. Upload `index.html`, `css`, `js`, and `data` to the repository root.
2. Open repository **Settings > Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Select branch `main` and folder `/(root)`, then save.

## Test locally
Opening `index.html` directly may block `fetch()` because of browser security. Start a local static server instead:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Notes
- The browser loads SheetJS from its official CDN, so users need internet access.
- The workbook and answers are downloadable by visitors. Do not use this design for confidential or formal examinations.
- GitHub Pages is static; scores are not stored centrally.

## If you see “Failed to fetch”

Do not double-click `index.html`. A `file:///` page cannot automatically fetch the adjacent workbook. Either publish with GitHub Pages, run `python -m http.server 8000`, or use the **Choose Excel file** fallback shown on the error screen. On GitHub, verify the exact case-sensitive path is `data/quiz_questions.xlsx`.
