# User guide

This guide covers the full translation workflow for CorelDRAW documents: exporting text, translating it, and importing the translations back.

- [The workflow at a glance](#the-workflow-at-a-glance)
- [Step 1: Export](#step-1-export)
- [Step 2: Translate](#step-2-translate)
- [Step 3: Import](#step-3-import)
- [Reading the import report](#reading-the-import-report)
- [Do's and don'ts](#dos-and-donts)
- [FAQ](#faq)

## The workflow at a glance

| Who | Step | Result |
|---|---|---|
| Designer / PM | Run **TranslationExport** on the final source document | One `.xlf` per target language, document saved with text IDs |
| Translator | Translate the `.xlf` in a CAT tool or TMS | Translated `.xlf` files |
| Designer / PM | Run **TranslationImport** on the same source document | One translated `.cdr` copy per language |
| Designer | Fix overflowing text and line breaks in each copy | Final translated documents |

The single most important rule: **don't edit the source document's text between export and import.** Text objects whose text changed are skipped on import.

## Step 1: Export

1. Open the source document. Make sure the text is final.
2. Run **TranslationExport** from the Scripts docker (**Tools › Scripts › Scripts**), your toolbar button, or **Alt+F8**.
3. Answer the three questions:

| Question | Meaning |
|---|---|
| **Source language** | Language of the document, as a code such as `de-CH`, `en-GB`, `fr`. |
| **Target languages** | Comma-separated codes, for example `fr-CH, it-CH, en-GB`. One XLIFF file is created per language. Leave empty to create a single file without a target language. |
| **Include text on hidden layers?** | Also export text on layers that are switched off. Default: No. |

4. The files are written next to the `.cdr` file, and the document is saved so the text IDs are kept:

```
Poster.cdr
Poster_fr-CH.xlf
Poster_it-CH.xlf
Poster_en-GB.xlf
```

A summary shows how many text objects and paragraphs were exported, and the folder opens in Explorer.

**What gets exported:** artistic text and paragraph text, on all pages, master layers and the desktop (pasteboard), including text inside groups and PowerClips and text fitted to a path, in reading order (page, then top to bottom, then left to right). Linked paragraph frames are exported once, as one story. Each paragraph becomes one translation unit.

**What doesn't:** text converted to curves, text inside symbols, bitmaps, and placed or linked files (PDF, AI, images).

> If the document has never been saved, the macro asks for a folder and reminds you to save the document. Unsaved IDs are lost when you close it.

## Step 2: Translate

Send the `.xlf` files to translation, or import them into your CAT tool. XLIFF 1.2 is supported by memoQ, Trados Studio, Phrase, Crowdin, Smartcat, MateCat and most other tools.

Instructions for translators:

- **Keep the inline tags.** `<g>` tags mark formatted words (bold, a different font, size or colour). Place them around the equivalent words in the translation.
- **Keep object tags in their order.** `<x ctype="x-coreldraw-object"/>` stands for an object embedded in the text. If one is deleted or two are swapped, that paragraph keeps its source text and is reported. Most CorelDRAW documents don't contain any.
- **Line breaks and tabs are flexible.** `<x ctype="lb"/>` (line break) and `<x ctype="x-tab"/>` (tab) can be moved, added or removed where the translation needs it.
- **Read the notes.** Each text object has a note with its page, layer, font and size. Objects marked *fixed text frame, watch the length* have limited room: aim for a translation of similar length.
- **Don't merge or split segments** across trans-units. Each unit is one paragraph in the layout.
- **Return the files under any name.** The target language is read from the file, and the file name is used as a fallback.

## Step 3: Import

1. Open the **original source document** (the one you exported from, saved with the IDs).
2. Run **TranslationImport**.
3. Select a translated `.xlf` file. If more XLIFF files for the same document are in that folder (for example `Poster_it-CH.xlf` next to `Poster_fr-CH.xlf`), you're asked whether to import them too.
4. A summary lists each file with its language and paragraph count, with a warning if a file was exported from a different document.
5. With a single file you choose:

| Answer | Meaning |
|---|---|
| **Yes** | Default. Creates `Poster_fr-CH.cdr` next to the original. The original is not modified. |
| **No** | Writes the translation into the open document. The whole import is one step in **Edit › Undo**. |

With several files, one translated copy per language is always created.

Each translated copy is created, filled, saved and left open. If a copy already exists, you're asked whether to overwrite it. Close any open copy before re-importing.

## Reading the import report

After the import, a summary appears per language:

```
fr-CH -> Poster_fr-CH.cdr
  Translated paragraphs: 58
  Not translated (original kept): 2
  Skipped, text changed since export: "Sommeraktion bis 31. August"
  Text objects not found (deleted?): 1
  OVERFLOWING TEXT in 1 frame(s): Page 1 "Découvrez notre nouvelle coll..."
```

| Line | What it means | What to do |
|---|---|---|
| **Translated paragraphs** | Paragraphs written successfully. | Nothing. |
| **Not translated** | The XLIFF had no target for these paragraphs. The source text was kept. | Check with the translator, or translate manually. |
| **Object placeholders missing or moved** | An object tag was deleted or reordered in the translation, so the paragraph was left in the source language. | Translate the paragraph manually, or have the translator restore the tags. |
| **Skipped, text changed since export** | The text in the document no longer matches the exported source, so the translation was not written. | Translate these objects manually, or re-export and send the changes again. |
| **Not found** | A text object with this ID no longer exists, or you opened a different document. | Make sure you opened the original source document. |
| **Duplicate IDs** | A text object was copied after export, so two objects share an ID. Only the first receives the translation. | Translate the copy manually. |
| **OVERFLOWING TEXT** | The translated text doesn't fit the paragraph frame. | Enlarge the frame, reduce the size or shorten the translation. |
| **ERRORS** | CorelDRAW refused a change for this object. Its original text was kept. | Translate it manually and report the error message. |

## Do's and don'ts

**Do**

- Export from the final, approved source document.
- Keep the source document (saved with IDs) until all languages are imported.
- Review every translated copy: line breaks, hyphenation and overflowing text always need a designer's eye.
- Re-run the export after larger source changes. Existing IDs are kept, so translation memory matches still work.

**Don't**

- Edit, add or delete text in the source between export and import.
- Run the import on a translated copy. Always import into the source.
- Convert text to curves before exporting.
- Rename or delete the `_p0`, `_p1` … trans-unit IDs in the XLIFF.

## FAQ

**Does the export change my document?**
It adds an invisible ID to each exported text object and saves the document. Nothing visible changes.

**Can I export several documents at once?**
Not yet. Run the export per document.

**Are fonts changed for languages with special characters?**
No. The import keeps the original font. If a font lacks glyphs for a language (for example Polish or Czech characters), check those copies carefully.

**Right-to-left languages (Arabic, Hebrew)?**
Text is written in, but proper display needs right-to-left text settings. Expect manual layout work.

**Does it work on a Mac?**
No. CorelDRAW for Mac has no VBA. Use the Windows version for the export and import; the `.cdr` files themselves can be edited on either platform.

**What if the translator returns XLIFF 2.0?**
Only XLIFF 1.2 is supported. Most tools can export 1.2 if you set the format when you create the project.
