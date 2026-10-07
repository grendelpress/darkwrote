# Darkwrote

A themed, in-browser editor for Word (`.docx`) files. Open a document, edit it with the formatting you
expect, and save it back — in a dark or light theme of your choice (or one you design yourself).
Everything runs locally in your browser; files are never uploaded.

## Features

- **Opens `.docx`** via Open, drag & drop, or Ctrl+O.
- **Displays Word formatting**: headings, bold / italic / underline / strike, sub/superscript, font family & size,
  text colour, highlight, alignment, bulleted & numbered (nested) lists, tables (merged cells, shading), images,
  hyperlinks, page size and margins.
- **Saves `.docx`**: Ctrl+S writes back to the opened file in Chromium browsers (File System Access API);
  elsewhere it downloads a copy.
- **Themes**: 6 dark + 4 light built in. *Customize…* lets you duplicate any theme, edit every colour with live
  preview, and export/import themes as JSON. Custom themes are stored in your browser.
- **Readable on dark themes**: colours authored in the document (e.g. black text, white table shading) are shifted
  into a readable range at display time only; the saved file keeps the original colours. Toggle this in
  *Customize…* → "Adapt document colours to theme".

## Development

```bash
npm install
npm run dev     # start the dev server
npm test        # import/export tests
npm run build   # type-check and build to dist/
```

## Known limitations

Round-tripping is done by Darkwrote's own reader/writer, so features it doesn't model are dropped on save:
headers/footers, footnotes, comments, tracked changes, text boxes/shapes, columns, custom paragraph spacing and
indents, and style definitions (formatting is saved as direct formatting). Keep a backup of important files.
Legacy `.doc` is not supported.
