# Zotero Reference Checker

## Overview

Zotero Reference Checker is a portable Windows tool for checking Microsoft Word documents that use live Zotero citations.

It creates one checked copy of the original document:

```text
OriginalFile_checked.docx
```

The original document is not modified.

---
## Requirements

- Windows
- Microsoft Word desktop app
- Zotero Desktop
- Synced Zotero group libraries
- Live Zotero citations in the Word document

The tool does not require administrator rights.

---

## Build Command

To build the portable executable:

```bash
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --onefile --noconsole ReferenceChecker.py
```

The executable will be created in:

```text
dist/ReferenceChecker.exe
```

## What the Tool Checks

The checker reviews:

1. Live Zotero in-text citations
2. Zotero-generated bibliography entries
3. Whether citations come from approved Zotero libraries
4. Broken in-text citations Zotero field codes 
5. Whether bibliography items are used in the text
6. Whether bibliography items have a DOI or URL
7. Whether bibliography items are duplicated 
8. Possible manually typed in-text citations

---

## Approved Zotero Libraries

The checker allows citations from only two Zotero group libraries:

1. `WGI AR7 General`
2. One additional Zotero group library selected by the user

---
## Input file

It is better to store the input file in a local folder rather than in cloud storage such as OneDrive.

---
## Recommended Workflow

Before running the checker:

1. Open Zotero.
2. Sync Zotero completely.
3. Make sure both required group libraries are synced.
4. Close Zotero.
5. Save and close the Word document.
6. Close all Microsoft Word windows.
7. Run `ReferenceChecker.exe`.
8. Select the second Zotero group library.
9. Select the Word document.
10. Wait for the finishing message.
11. Open `OriginalFile_checked.docx`.

Do not keep Zotero or the target Word document open while the checker runs.

---

## Output

The tool creates:

```text
OriginalFile_checked.docx
```
This checked file contains:

- Word comments
- issue descriptions directly in the document

It also creates or updates:

``` text
refcheck_debug.txt
```
in the same folder as the selected Word document.

The debug log records essential information only for issues that actually receive Word comments. Each run includes the date and time, input file, and selected library.

---

### Linked Zotero Citation Issues

#### "This appears to be a damaged Zotero citation field."

The text is still inside a Zotero-backed Word field, but the field code is
not a valid `ZOTERO_ITEM` citation field—for example, a damaged code such
as `ADDIN Zotero TEMPP`.

The citation may have become corrupted or partially unlinked. Reinsert
the citation using Zotero. Damaged Zotero fields are excluded from the
manual-citation check, so they should receive this comment instead of a
second "manually typed citation" comment.

<ins>For this type of comment, try to find out if the in-text citations can still be found in the reference list. If yes, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

#### "This is a linked Zotero citation, but the checker could not read its Zotero item data."

The Word field looks like a Zotero citation, but its embedded citation
data could not be parsed. The field may be damaged, incomplete, or
partially copied from another document.

<ins>For this type of comment, check if it is really an in-text citation. If not, delete the comment.</ins> 

#### "The checker could not read the item key for one citation item."

The Zotero field data was readable, but one citation item did not contain
a usable Zotero item key. The field may be incomplete or damaged.

<ins>For this type of comment, try to find out if the in-text citations can still be found in the reference list. If yes, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

#### "This linked Zotero citation item was not found in the selected local Zotero libraries."

The exact item key embedded in the Word citation was not found in either
approved library. Before adding this comment, the checker also searches
the selected libraries for an item with the same normalized embedded
title.

This comment therefore means that neither:

- the embedded Zotero item key, nor
- an exact normalized title match

was found in the two selected libraries.

Possible causes include:

- the original Zotero item was deleted, replaced, merged, or substantially modified
- the citation was copied from another document or library
- the relevant group library is not fully synced
- the wrong second group library was selected
- the selected-library copy has a materially different title

<ins>For this type of comment, try to find out if the in-text citations can still be found in the reference list. If yes, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

#### "Citation item comes from unexpected Zotero library."

The citation resolves to an item outside:

- `WGI AR7 General`
- the selected second group library

This warning is mainly a safeguard. Bibliography and title matching are
otherwise restricted to the two selected libraries.

<ins>For this type of comment, try to find out if the in-text citations can still be found in the reference list. If yes, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

------------------------------------------------------------------------

### Bibliography Issues

#### "This bibliography entry appears in the reference list but was not found as a linked in-text citation."

The checker matched the bibliography entry to a Zotero item, but could
not connect it to any linked in-text citation using:

- the Zotero item key
- the normalized item title
- another cited title contained within the same bibliography entry

The final safeguard is important for chapters, annexes, glossary entries,
and similar references that also contain a parent book or report title.

Title comparison normalizes:

- HTML formatting tags
- HTML entities such as `&thinsp;`
- straight and curly quotation marks
- Unicode hyphens and dashes
- Unicode subscript digits such as `SO₂` and `SO2`
- punctuation, capitalization, and spacing

A remaining comment may indicate that the original cited Zotero item was
deleted or modified, or that the bibliography entry is genuinely not
cited.

<ins>For this type of comment, try to find out if an in-text citation corresponding to the bibliography really cannot be found in the text. If yes, delete the comment. Otherwise, delete the bibliography through the Zotero add-on.</ins>

#### "This bibliography entry duplicates another entry in the reference list and matches the same Zotero item."

The same normalized bibliography entry appears more than once and both
copies match the same Zotero item key. The later occurrence receives the
comment.

<ins>For this type of comment, try to find out if it is true. If not true, delete the comment. Otherwise, delete the duplication through the Zotero add-on.</ins>


#### "This bibliography entry has the same title as another entry but matches a different Zotero item."

Two identical normalized bibliography entries have the same matched
title but different Zotero item keys. This may indicate duplicate Zotero
records in the selected libraries.

The checker requires the normalized full bibliography entries to match,
which reduces false duplicate warnings for different chapters that share
the same parent book or report title.

<ins>For this type of comment, try to find out if it is true. If not true, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

#### "This reference has neither DOI nor URL in Zotero."

This comment is added only when all four checks are negative:

- the matched Zotero item's DOI field is empty
- the matched Zotero item's URL field is empty
- the visible bibliography entry contains no DOI text
- the visible bibliography entry contains no `http://` or `https://` URL

A visible DOI or URL therefore prevents this comment even when the
corresponding Zotero metadata fields are empty.

<ins>For this type of comment, try to find out if the article can be found online. If not, try to find a Zotero entry (with an attached PDF) in the "Non-peer-reviewed literature" folder of the "WGI AR7 General" library. If an entry can be found, delete the comment. If not, reply to the comment with "Please provide the metadata and the PDF file of the article in the 'Non-peer-reviewed literature' folder."

------------------------------------------------------------------------

### Possible Manual Citation Issues

#### "This appears to be a manually typed in-text citation."

The checker found author–year citation-like text that does not overlap a
valid or damaged Zotero-backed Word field.

Example forms checked include:

```text
(Gong et al., 2020)
Gong et al. (2020)
Jenkins et al. (2022)
Chen, W et al. (2002)
Li and Paul (2026)

<ins>For this type of comment, try to find out if the in-text citations can still be found in the reference list. If yes, delete the comment. Otherwise, keep it for the reference PoCs to re-insert the in-text citations.</ins>

---
## Debug Log

`refcheck_debug.txt` is saved in the same folder as the Word document
being checked.

The log no longer records every in-text citation field. It records
concise diagnostic information only when the checker adds a Word
comment, making it easier to investigate false positives.

The log is appended across runs and separates each run with a timestamp
and run information.

---

## What Is Ignored

The manual citation checker ignores date ranges and year ranges such as:

```text
(1850–1900)
(2004–2023)
(2026–2035)
(2020/2021)
(2020 to 2030)
```

It also ignores references to figures, tables, sections, equations, boxes, annexes, and supplements.

---


---

## Important Limitations

The checker is most reliable for live Zotero citations.

Possible manually typed citation detection is pattern-based. It may require human review.

The checker does not edit the original document.
