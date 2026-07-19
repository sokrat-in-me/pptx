# AGENTS.md

## Cursor Cloud specific instructions

This repo is a single Python 3 CLI utility (`generate_presentation.py`) that reads
`data/RTR-status.csv` and produces a PowerPoint deck. There are no servers,
databases, or web services to start.

- Run: `python3 generate_presentation.py` (see `README.md`). Output defaults to
  `Дефекты_и_разработки_по_Пилоту.pptx`. Override with `--csv` / `--output`.
- No lint or automated test tooling is configured. Validate changes by running the
  generator and inspecting the `.pptx` (e.g. open it, or convert with LibreOffice).
  A syntax check is `python3 -m py_compile generate_presentation.py`.
- Regenerating the deck rewrites the committed `.pptx` with a byte-different file
  (zip embeds timestamps) even when content is unchanged, which dirties the git
  tree. Run `git checkout -- '<output>.pptx'` if you did not intend to change it.
- Row filtering is branch-specific: this branch only includes CSV rows containing
  «Осипова» and not «важно» (see `generate_presentation.py`). Related generators
  live on other branches.
- To render the deck to images for visual verification (not a project dependency),
  install `libreoffice-impress` + `poppler-utils`, then
  `soffice --headless --convert-to pdf --outdir <dir> <file>.pptx` and
  `pdftoppm -png <file>.pdf slide`.
