# Changelog

All notable changes to this project are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-10-05

### Added

- `SupertextTranslation.bas` with two macros:
  - `TranslationExport`: exports artistic and paragraph text to XLIFF 1.2, one file per target language, with persistent `SupertextID` shape properties, paragraph-level trans-units, `<g>` inline formatting tags, `<x>` tags for line breaks, tabs and embedded objects, and context notes.
  - `TranslationImport`: imports translated XLIFF into per-language copies as one undo step, copies the original character formatting onto the translated words, skips text changed since export, rejects missing or reordered object tags, unlocks locked layers and objects, and reports overflowing frames.
- `tests/apply-paragraph-test.js`: tests the import's paragraph logic through a JavaScript port.
- Installation guide, user guide and developer guide.

### Known issues

- Not yet compiled or tested in CorelDRAW. See "Assumptions to verify in CorelDRAW" and the test checklist in the developer guide.
