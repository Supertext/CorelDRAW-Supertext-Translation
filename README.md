# CorelDRAW Supertext Translation

Round-trip translation for CorelDRAW documents, by [Supertext](https://www.supertext.com).

Export all text from a `.cdr` document to **XLIFF 1.2**, translate it in any CAT tool or TMS (memoQ, Trados, Phrase, …), and import the translations back into the **same text objects with their formatting intact**: one translated copy of the document per language.

```
Poster.cdr ──► TranslationExport ──► Poster_fr-CH.xlf ──► CAT tool / TMS
                                     Poster_it-CH.xlf          │
                                                               ▼
Poster_fr-CH.cdr ◄── TranslationImport ◄── translated .xlf files
Poster_it-CH.cdr
```

The XLIFF format matches [Adobe Illustrator Translation](https://github.com/Supertext/Adobe-Illustrator-Translation) and [Adobe InDesign Translation](https://github.com/Supertext/Adobe-InDesign-Translation), so one translation workflow covers all three applications.

> **Status: 1.0.0, untested in CorelDRAW.** The macros have been reviewed by hand and the paragraph replacement logic passes a test against a JavaScript port, but they have not yet been compiled or run in CorelDRAW. Run the [test checklist](docs/developer-guide.md#test-checklist) on a sample file before using this in production.

## Features

- **Stable text IDs.** Each text object gets a hidden ID stored inside the `.cdr` file, so translations land in the right object even if it is moved.
- **Artistic and paragraph text**, on all pages, master layers and the desktop, including text in groups and PowerClips. Linked paragraph frames are exported once per story.
- **Formatting preserved.** One translation unit per paragraph. A bold word, a different font, size or colour inside a paragraph becomes an XLIFF `<g>` tag. On import the translated words take over the formatting of the original characters, so fills and outlines carry over.
- **Context for translators.** Every text object carries a `<note>` with its page, layer, font, size and text type, and paragraph frames are flagged "watch the length".
- **Safe import.** The original stays untouched. Text edited after export is skipped and reported. Untranslated paragraphs keep the source text. Locked objects and layers are handled automatically. Each import is a single undo step.
- **Overflow check.** After import, paragraph frames whose translated text no longer fits are listed.

## Quick start

1. Import [`src/SupertextTranslation.bas`](src/SupertextTranslation.bas) into CorelDRAW's VBA editor ([installation guide](docs/installation-guide.md)).
2. Open your document and run the macro **TranslationExport**. Enter the source and target languages.
3. Translate the `.xlf` files.
4. Open the original document again and run **TranslationImport**. Select a translated `.xlf` file.

## Documentation

| Guide | For |
|---|---|
| [Installation guide](docs/installation-guide.md) | Installing, updating and removing the macros |
| [User guide](docs/user-guide.md) | Designers, project managers and translators running the workflow |
| [Developer guide](docs/developer-guide.md) | Architecture, XLIFF format, testing and releasing |

## Repository layout

```
src/
  SupertextTranslation.bas     VBA module with the TranslationExport and TranslationImport macros
tests/
  apply-paragraph-test.js      tests the import's paragraph logic via a JavaScript port (Node.js)
docs/
  installation-guide.md
  user-guide.md
  developer-guide.md
CHANGELOG.md
```

## Requirements

- CorelDRAW Graphics Suite 2020 or later **on Windows**, with Visual Basic for Applications installed (an option in the CorelDRAW installer).
- CorelDRAW for Mac is not supported, because it has no VBA.
