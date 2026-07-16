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
4. Broken in-text citations
5. Whether bibliography items are used in the text
6. Whether bibliography items have a DOI or URL
7. Whether bibliography items are duplicated 
8. Possible manually typed in-text citations

---

## Approved Zotero Libraries

The checker allows citations from only two Zotero group libraries:

1. `WGI AR7 General`
2. One additional Zotero group library selected by the user

Items from the personal Zotero library or other group libraries are commented as unexpected-library issues.

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

## Meaning of Comments

### Linked Zotero Citation Issues

#### “This is a linked Zotero citation, but the checker could not read its Zotero item data.”

The citation is linked, but its Zotero field metadata may be damaged or incomplete.

#### “This linked Zotero citation item was not found in the local Zotero database.”

The citation points to a Zotero item that was not found locally.

Possible causes:

- Zotero library not synced
- item deleted
- citation copied from another document
- broken Zotero field
- wrong second group library selected

#### “Citation item comes from unexpected Zotero library.”

The citation is linked, but the item belongs to a library other than:

- `WGI AR7 General`
- the selected second group library

---

### Bibliography Issues

#### “This bibliography entry appears in the reference list but was not found as a linked in-text citation.”

The reference appears in the bibliography but was not found among linked Zotero citations.

The checker uses item keys and normalized titles. It also checks whether a bibliography entry contains a cited chapter or annex title when the same entry also contains a parent book or report title.


#### “This reference has neither DOI nor URL in Zotero.”

The Zotero item has no DOI and no URL.

#### “This bibliography item comes from unexpected Zotero library.”

The bibliography item belongs to a Zotero library that is not approved for this check.

---

### Possible Manual Citation Issues

#### “This appears to be a manually typed in-text citation.”

The checker found citation-like text that does not appear to be backed by a Zotero field.

Example forms checked include:

```text
(Gong et al., 2020)
Gong et al. (2020)
Jenkins et al. (2022)
Chen, W et al. (2002)
Li and Paul (2026)
```

The tool adds a comment only.
The comment asks the user to insert the citation using Zotero and update the reference list.

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
