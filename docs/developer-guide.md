# Developer guide

How the macros work, what the XLIFF looks like, and how to test and release changes.

- [Overview](#overview)
- [VBA constraints](#vba-constraints)
- [Text IDs](#text-ids)
- [Text model: stories, paragraphs, style runs and objects](#text-model-stories-paragraphs-style-runs-and-objects)
- [XLIFF format](#xliff-format)
- [Import algorithm](#import-algorithm)
- [Assumptions to verify in CorelDRAW](#assumptions-to-verify-in-coreldraw)
- [Development setup](#development-setup)
- [Test checklist](#test-checklist)
- [Releasing](#releasing)
- [Known limitations and roadmap](#known-limitations-and-roadmap)

## Overview

| File | Role |
|---|---|
| `src/SupertextTranslation.bas` | One VBA module. Public macros `TranslationExport` and `TranslationImport`; everything else is private. |
| `tests/apply-paragraph-test.js` | A JavaScript port of `ApplyParagraph`, tested against a mock story in Node.js |

Export and import live in one module, so the code they share (shape collection, IDs, text analysis, run detection) exists only once. That's a difference from the Illustrator and InDesign scripts, which must stay self-contained per file.

This project mirrors [Adobe Illustrator Translation](https://github.com/Supertext/Adobe-Illustrator-Translation) and [Adobe InDesign Translation](https://github.com/Supertext/Adobe-InDesign-Translation). The XLIFF structure, ID conventions and import report are deliberately the same.

## VBA constraints

- **Windows only.** CorelDRAW for Mac has no VBA. A cross-platform version would need CorelDRAW's JavaScript scripting (see the roadmap).
- **Late binding** for everything outside CorelDRAW: `Scripting.Dictionary`, `Scripting.FileSystemObject`, `MSXML2.DOMDocument.6.0`, `ADODB.Stream`, `VBScript.RegExp`, `Shell.Application` are all created with `CreateObject`, so no references need to be set in the VBA project. All of them ship with Windows.
- **Unicode:** the code uses `WideText`, `InsertBeforeWide` and `InsertAfterWide` and builds special characters with `ChrW`. XLIFF is written as UTF-8 with `ADODB.Stream` (with a BOM, which CAT tools accept) and parsed by MSXML. Keep the `.bas` source itself ASCII: the VBA editor reads imported files in the system code page.
- **CRLF line endings.** The VBA editor's import requires Windows line endings. `.gitattributes` enforces `eol=crlf` for `*.bas`.
- **ByVal parameters.** Most parameters are declared `ByVal`. Passing an element of a Variant array or Dictionary to a `ByRef` parameter of a specific type is a compile error ("ByRef argument type mismatch").
- **No user forms.** Dialogs use `InputBox` and `MsgBox`, so the module is one plain-text file without a binary `.frx`. A proper dialog is on the roadmap.

## Text IDs

The ID is stored as persistent shape data: `shape.Properties("SupertextID", 1)`. It is saved in the `.cdr`, invisible in the UI, and survives moving and restyling. For linked paragraph frames it is stored on the first frame, which carries the story.

```
id = "cd" + hex(timer) + hex(random)   ' e.g. cd4f1a2b7c3e9d01
```

`EnsureId` rules:

- An existing ID is reused, so re-exporting keeps the same IDs and TM matches stay stable.
- If two shapes carry the same ID (a shape was copied), the first keeps it and the later one gets a new ID on export. On import, duplicates are counted and only the first is used.
- Writing the ID needs an editable layer and an unlocked shape. `WithUnlocked` makes the layer editable and unlocks the shape, then restores both.

**Removing IDs** (for example before handing a file to a client): add this to any module and run it once.

```vb
Sub RemoveSupertextIDs()
    Dim p As Page, l As Layer, s As Shape
    For Each p In ActiveDocument.Pages
        For Each l In p.Layers
            For Each s In l.Shapes.FindShapes(Type:=cdrTextShape)
                On Error Resume Next
                s.Properties.Delete "SupertextID", 1
                On Error GoTo 0
            Next
        Next
    Next
End Sub
```

## Text model: stories, paragraphs, style runs and objects

`CollectTextShapes` walks every page's layers and then the master page's layers (master layers and the desktop), recursing into groups and PowerClips. Shapes are deduplicated by `StaticID`. For paragraph text, only the first frame of a linked chain is processed (`IsFirstFrame`), and `Shape.Text.Story` gives the text of all linked frames.

`Analyze(shape)` returns:

```
{ text, paras: [ { start, text, runs: [ Array(start, len, key) ] } ] }
```

- **Story text** comes from `Story.WideText`. Positions are character positions in the story. If the string length and `Story.Length` disagree, `CRLF` pairs are treated as one paragraph break; if they still disagree, the object is skipped and reported, rather than writing to the wrong positions.
- **Paragraphs** are split on `vbCr`. **Line breaks** are `vbLf` or `Chr(11)`; both are exported as `<x ctype="lb"/>`. On import, a paragraph that already contains a line break reuses that character; otherwise `DEFAULT_LINE_BREAK` is used.
- **Style runs:** `StyleKey` combines font, size, bold, italic, underline, strikethrough, position, case, character spacing and fill type and colour. Adjacent ranges with the same key are merged.
- **Objects:** any control character other than tab, line break and paragraph break, plus `U+FFFC` and `U+FEFF`, is treated as an embedded object and exported as a protected placeholder (`IsHard`).

### Run detection modes

`RUN_DETECTION` at the top of the module:

| Mode | How | Speed |
|---|---|---|
| `"fast"` (default) | `Story.EnumRanges()` returns ranges of identical character attributes. Falls back to `exact` if their lengths don't add up to the story length. | Fast |
| `"exact"` | Computes `StyleKey` for every character. | Slow on long texts |

### Range positions

All position-based access goes through one function, `Span(shape, pos, length)`, with 0-based positions. CorelDRAW's documentation doesn't state whether `TextRange.Range(Start, End)` is 0-based with an exclusive end, so `CalibrateRange` tests both conventions on the first story with at least two characters and remembers the one that returns the expected characters.

## XLIFF format

XLIFF 1.2, one `<file>` per document and target language, one `<group>` per text object (story), one `<trans-unit>` per paragraph with translatable text.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file original="Poster.cdr" source-language="de-CH" target-language="fr-CH" datatype="x-coreldraw">
    <header>
      <tool tool-id="supertext-coreldraw" tool-name="Supertext CorelDRAW Translation" tool-version="1.0.0"/>
    </header>
    <body>
      <!-- group id = the shape's SupertextID -->
      <group id="cd4f1a2b7c3e9d01">
        <note>Page 1 | Layer: Text | paragraph text | Arial 24 pt | fixed text frame, watch the length</note>
        <!-- trans-unit id = <group id>_p<paragraph index> -->
        <trans-unit id="cd4f1a2b7c3e9d01_p0" xml:space="preserve">
          <!-- run 0 is the paragraph's base style; <g id="1"> = style of run 1 -->
          <source>Jetzt <g id="1">neu</g> im Sortiment<x id="lb1" ctype="lb"/>ab 1. März</source>
          <target>Désormais <g id="1">nouveau</g> dans notre assortiment<x id="lb1" ctype="lb"/>dès le 1er mars</target>
        </trans-unit>
      </group>
    </body>
  </file>
</xliff>
```

Conventions (shared with the Illustrator and InDesign versions):

- **Paragraph index** `_pN` is the index in the story's `vbCr`-split, including empty paragraphs. Paragraphs without translatable text get no unit, so indices can have gaps.
- **`<g id="k">`** wraps text whose style differs from run 0. `k` is the 0-based run index within the paragraph. Text in run 0's style is emitted untagged.
- **`<x ctype="lb"/>`** and **`<x ctype="x-tab"/>`** are soft: translators may move, add or drop them. Literal newlines in a target are treated as line breaks.
- **`<x ctype="x-coreldraw-object"/>`** with `id="o1"`, `o2`, … numbered per paragraph is a hard placeholder: targets must contain all of them in the same order.
- **`<mrk>` and `<sub>`** in targets are unwrapped, other unknown elements ignored.
- **Target language** comes from `target-language`, falling back to a `_xx-YY` suffix in the file name.

## Import algorithm

Per XLIFF file (`ApplyPack`), inside `Document.BeginCommandGroup` / `EndCommandGroup` (one undo step) with `Optimization` on:

1. Index the document's text shapes by ID.
2. For each group in the XLIFF (`ApplyShape`), with layer and shape unlocked, and errors caught per shape (`TryApply`) so one failing object doesn't stop the import:
   1. **Verify.** Re-analyse the story. If any unit's `<source>` (flattened, line breaks as LF, objects as U+E000) differs from the current paragraph text, skip it as *changed*.
   2. **Process paragraphs from last to first**, so positions of earlier paragraphs stay valid. Untranslated paragraphs are not touched.
3. Per translated paragraph (`ApplyParagraph`):
   1. Split the target at its object placeholders. If they aren't exactly `o1…oN` in order, leave the paragraph unchanged and report it.
   2. The paragraph is cut into **chunks** between object characters.
   3. **Pass 1**, last chunk first: insert the chunk's translation before its old text (`InsertBeforeWide`), then give each segment the formatting of the first character of its original style run (`CopyAttributes`). `loc(k)` tracks where run *k*'s first character currently is as insertions shift the text. All original characters still exist during this pass, so a translator may move a formatted word anywhere in the paragraph.
   4. **Pass 2**, last chunk first: delete each chunk's old text.

   Object characters and paragraph breaks are never deleted, so embedded objects and paragraph formatting survive. Copying attributes from the original characters means fills, outlines and any other character formatting carry over, not just the properties in `StyleKey`.
4. Check overflow with `Shape.Text.Overflow` for paragraph text.
5. Save the copy and collect the report.

## Assumptions to verify in CorelDRAW

The code was written against the CorelDRAW object model documentation without access to CorelDRAW. These points are the most likely to need adjustment in the first test, roughly in order of risk:

| Assumption | Where | If it's wrong |
|---|---|---|
| `TextRange.CopyAttributes SourceRange` copies character formatting with one argument | `ApplyParagraph` | Translated text keeps the formatting of the insertion point. Check the method's optional parameters. |
| `TextRange.Range(Start, End)` follows one of the two conventions `CalibrateRange` tests | `Span` | Text lands at wrong positions. Adjust `Span`. |
| `Shape.Properties(Name, Index)` stores hidden persistent data | `GetId`, `EnsureId` | IDs aren't found on import ("not found"). An alternative is `Shape.ObjectData`. |
| Line breaks in `WideText` are `vbLf` or `Chr(11)` | `IsLineBreak`, `DEFAULT_LINE_BREAK` | Line breaks are lost or become paragraph breaks. |
| `Shape.Text.Frame.Range.Start` is the story start for the first linked frame | `IsFirstFrame` | Linked text is exported twice or not at all. |
| `CorelScriptTools.GetFileBox` shows a file dialog | `PickXliffFiles` | The macro falls back to asking for the path in an input box. |
| `Story.EnumRanges()` lengths add up to `Story.Length` | `FastRuns` | The macro falls back to the slow exact mode; no error. |

## Development setup

- **Source of truth is `src/SupertextTranslation.bas`.** Edit it in the VBA editor (**Alt+F11**) for IntelliSense and debugging, then **File › Export File…** back over the repository file before committing. Or edit in a text editor and re-import (remove the old module first).
- **Compile** with **Debug › Compile** after every change. VBA reports many errors only at compile time.
- **Debugging:** set breakpoints (F9) in the VBA editor, run the macro from the Scripts docker, and inspect variables in the **Locals** window. `Debug.Print` writes to the **Immediate** window (Ctrl+G).
- **Logic test without CorelDRAW:** `node tests/apply-paragraph-test.js`. It contains a line-by-line JavaScript port of `ApplyParagraph` and checks text, formatting of moved words (including a word moved across an object), and placeholder validation against a mock story. When you change `ApplyParagraph`, change the port too.

## Test checklist

Build a test document `test/roundtrip.cdr` (not committed) containing:

- [ ] Artistic text, paragraph text, text fitted to a path
- [ ] Paragraph text linked across two frames on two pages
- [ ] A paragraph with a **bold** word, a word in another font and a coloured word, and one with an outline
- [ ] A line break (Shift+Enter), a tab, an empty paragraph
- [ ] Text in a group, inside a PowerClip, on a master layer, on the desktop
- [ ] Text on a locked layer, a locked object, text on a hidden layer
- [ ] Umlauts and special characters: `äöü ÄÖÜ ß é è à ç € « » – “ ” & <tag>`
- [ ] Left, centred and right aligned paragraphs

Round trip:

1. Export with targets `fr-CH, it-CH`. Open the XLIFF in a validator or CAT tool: it must parse, with the expected unit count. Check that umlauts are correct.
2. **Identity test:** copy every `<source>` into a `<target>` and import. The result must be visually identical to the source. This is the main test for the [assumptions](#assumptions-to-verify-in-coreldraw).
3. **Real test:** translate with longer text and moved `<g>` tags, then import. Check formatting lands on the right words, line breaks stay line breaks, and overflowing frames are reported.
4. **Safety tests:** edit one text before import (reported as changed), delete one (not found), copy one after export (duplicate ID).
5. **Undo:** with "apply to the open document", one **Edit › Undo** must revert the whole import.
6. Test in the oldest and newest CorelDRAW version you support, in both run detection modes if time allows.

## Code quality and security checks

- **Checks** workflow (`.github/workflows/checks.yml`): [actionlint](https://github.com/rhysd/actionlint) and [zizmor](https://docs.zizmor.sh/) lint the workflows on every push and pull request. Dependency review fails a pull request that adds a package with a known vulnerability (moderate or worse). Actions are pinned to commit SHAs; Dependabot keeps the pins up to date. To run the workflow lint locally: `pip install actionlint-py zizmor`, then `actionlint` and `zizmor .github/workflows` in the repo root.
- **Links** workflow (`.github/workflows/links.yml`): [lychee](https://lychee.cli.rs/) checks the links in all Markdown files weekly and whenever docs change on `main`. Broken links open (or update) the issue "Broken links in the docs". Links that can't work from CI (local URLs, placeholders, pages behind a login) are excluded in `.lycheeignore`.
- There is no build or CI test run yet: run the test above before committing. CodeQL (see below) analyses the JavaScript test; it doesn't cover VBA, so review `src/SupertextTranslation.bas` by hand.
- GitHub settings (set by Remy's setup script, not stored in the repo): **secret scanning with push protection** (a push containing a known token format is rejected; findings under *Security → Secret scanning*) and **CodeQL default setup** (findings under *Security → Code scanning* and as comments on pull requests; PHP isn't covered by CodeQL, which is why the PHP plugins run PHPStan).

Before starting work in this repo, look at its open findings: code scanning alerts, secret scanning alerts, Dependabot PRs and the "Broken links in the docs" issue.

## Releasing

1. Update `VERSION` at the top of the module (written into the XLIFF `<tool>`).
2. Add an entry to `CHANGELOG.md`.
3. Run `node tests/apply-paragraph-test.js` and the test checklist.
4. Tag the release: `git tag v1.1.0 && git push --tags`. Attach `SupertextTranslation.bas` to a GitHub release, and optionally a built `SupertextTranslation.gms` for one-file rollout.

Versioning follows [SemVer](https://semver.org). A change that alters the XLIFF structure or the ID scheme is a **major** version.

## Known limitations and roadmap

**Limitations**

- Untested in CorelDRAW so far; see [Assumptions to verify](#assumptions-to-verify-in-coreldraw).
- Windows only.
- Text in symbols, bitmaps, placed files and text converted to curves is not exported.
- Only XLIFF 1.2 is supported. In an XLIFF with several `<file>` elements, the target language is taken from the first one.
- No batch export across many documents yet.

**Roadmap ideas**

- **Dialog and docker:** a proper form with language presets, or a docker with export/import buttons and a clickable list of overflowing frames.
- **Mac support** via CorelDRAW's JavaScript scripting, sharing the XLIFF format.
- **Supertext API integration:** send the XLIFF straight to Supertext and import on completion.
- **XLIFF 2.0** export and import.
- **Batch mode** for a folder of `.cdr` files.
