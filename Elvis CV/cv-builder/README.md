# CV builder

Generates `Elvis_Nosakhare_CV_Technical_Lead.docx` (the "IT Technical Lead & Full-Stack
Developer" CV variant) from code, so the CV's content and layout are versioned and
reproducible instead of living only inside a hand-edited Word file.

## Usage

```bash
cd "Elvis CV/cv-builder"
npm install
npm run build          # writes ../Elvis_Nosakhare_CV_Technical_Lead.docx
```

Or directly: `node build_cv.js path/to/output.docx`.

## Notes

- Built with [`docx`](https://www.npmjs.com/package/docx) (docx-js). No other dependencies.
- Styling mirrors the original hand-built `Elvis_Nosakhare_CV.docx`: Calibri, navy
  `#0C2340` headings, amber `#B5720A` accents/links, slate `#44515E` body text, US
  Letter page, tight margins (all in `build_cv.js`'s constants/helpers at the top).
- Deliberately kept to **2 pages** — that's why entries are compact (one line per
  skill/project + an italic stack line) rather than more spaced out. If you add
  content, check the page count before committing (see below).
- To change the CV, edit `build_cv.js` directly (plain JS — search for the section
  you want, e.g. `sectionHeader("PROJECTS")`) and rebuild.

## Checking the page count

This machine had no LibreOffice/pandoc, and Microsoft Word's AppleScript export
hung — so page count was verified with a throwaway HTML mirror of the CV rendered
through headless Chrome (`Google Chrome --headless --print-to-pdf`), with Calibri
embedded from Word's own font files
(`/Applications/Microsoft Word.app/Contents/Resources/DFonts`). That verification
script wasn't preserved in the repo (it lived in a Claude Code session scratchpad).
On a machine with Word or LibreOffice installed, just open the generated `.docx` and
check **File → Print** / the page count in the status bar instead.
