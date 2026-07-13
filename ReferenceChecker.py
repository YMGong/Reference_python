import os
import re
import html
import json
import shutil
import sqlite3
import tempfile
import urllib.parse
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import win32com.client as win32


DEFAULT_LIBRARY = "WGI AR7 General"
WD_FORMAT_XML_DOCUMENT = 12


# -----------------------------
# General helpers
# -----------------------------

def clean_path(path):
    return os.path.abspath(os.path.normpath(urllib.parse.unquote(path)))


def short_text(text, limit=100):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text[:limit] + ("..." if len(text) > limit else "")


def normalize_text(text):
    return re.sub(r"\s+", " ", text or "").strip().lower()

def normalize_match_text(text):
    text = html.unescape(text or "")
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.lower()

    # Normalize Unicode subscript digits, e.g. SO₂ -> SO2.
    text = text.translate(str.maketrans({
        "₀": "0",
        "₁": "1",
        "₂": "2",
        "₃": "3",
        "₄": "4",
        "₅": "5",
        "₆": "6",
        "₇": "7",
        "₈": "8",
        "₉": "9",
    }))

    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u00a0", " ")
    text = text.replace("\u2010", "-")
    text = text.replace("\u2011", "-")
    text = text.replace("\u2012", "-")
    text = text.replace("\u2013", "-")
    text = text.replace("\u2014", "-")
    text = text.replace("\u2212", "-")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# Debug logging
debug_path = None

def debug_log(text):
    if not debug_path:
        return

    with open(debug_path, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def start_debug_run(input_path, selected_library):
    with open(debug_path, "a", encoding="utf-8") as f:
        f.write("\n\n==============================\n")
        f.write("REFERENCE CHECK RUN: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        f.write("Input file: " + str(input_path) + "\n")
        f.write("Selected library: " + str(selected_library) + "\n")
        f.write("==============================\n")


def log_comment_context(issue_type, details):
    debug_log("\n--- COMMENT ADDED: " + issue_type + " ---")
    for key, value in details.items():
        debug_log(str(key) + ": " + repr(value))


def log_elapsed_time(check_name, started_at):
    """Write elapsed time to the debug log."""
    elapsed = time.perf_counter() - started_at
    debug_log(f"TIME - {check_name}: {elapsed:.3f} seconds")
    return elapsed


def normalize_citation_component(text):
    text = text.lower()
    text = re.sub(r"[().,;:]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def range_text(word_range):
    try:
        return short_text(word_range.Text)
    except Exception:
        return ""


def paragraph_anchor_range(paragraph):
    r = paragraph.Range.Duplicate
    try:
        if r.End > r.Start:
            r.End = r.End - 1
    except Exception:
        pass
    return r


def ranges_overlap(start1, end1, start2, end2):
    return start1 < end2 and start2 < end1


# -----------------------------
# Zotero database helpers
# -----------------------------

def find_zotero_db():
    home = os.path.expanduser("~")
    roots = [
        os.path.join(home, "Zotero"),
        os.path.join(home, "AppData", "Roaming", "Zotero"),
        os.path.join(home, "AppData", "Local", "Zotero"),
    ]

    for root in roots:
        if os.path.exists(root):
            for dirpath, _, files in os.walk(root):
                if "zotero.sqlite" in files:
                    return os.path.join(dirpath, "zotero.sqlite")
    return None


def choose_zotero_db():
    return filedialog.askopenfilename(
        title="Select zotero.sqlite",
        filetypes=[
            ("Zotero database", "zotero.sqlite"),
            ("SQLite files", "*.sqlite"),
            ("All files", "*.*"),
        ],
    )


def db_connect(db_path):
    return sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)


def get_table_columns(cur, table_name):
    try:
        cur.execute(f"PRAGMA table_info({table_name})")
        return [row[1] for row in cur.fetchall()]
    except Exception:
        return []


def load_library_names(db_path):
    conn = db_connect(db_path)
    cur = conn.cursor()
    libraries = {1: "Personal Zotero library (My Library)"}

    cols = get_table_columns(cur, "groups")

    if "libraryID" in cols:
        name_col = None
        for candidate in ["name", "groupName"]:
            if candidate in cols:
                name_col = candidate
                break

        if name_col:
            try:
                cur.execute(f"SELECT libraryID, {name_col} FROM groups")
                for library_id, name in cur.fetchall():
                    if name:
                        libraries[library_id] = name
            except Exception:
                pass

    conn.close()
    return libraries


def load_zotero_items(db_path):
    conn = db_connect(db_path)
    cur = conn.cursor()

    cur.execute("SELECT itemID, key, libraryID FROM items")
    rows = cur.fetchall()

    library_names = load_library_names(db_path)

    items = {}
    item_id_to_key = {}

    for item_id, key, library_id in rows:
        if not key:
            continue

        library_name = library_names.get(
            library_id,
            f"Unknown Zotero library ID {library_id}"
        )

        items[key] = {
            "itemID": item_id,
            "key": key,
            "libraryID": library_id,
            "library": library_name,
            "title": "",
            "DOI": "",
            "url": "",
        }

        item_id_to_key[item_id] = key

    try:
        cur.execute("""
            SELECT itemData.itemID, fields.fieldName, itemDataValues.value
            FROM itemData
            JOIN fields ON itemData.fieldID = fields.fieldID
            JOIN itemDataValues ON itemData.valueID = itemDataValues.valueID
        """)

        for item_id, field_name, value in cur.fetchall():
            key = item_id_to_key.get(item_id)
            if not key:
                continue

            field_lower = field_name.lower()

            if field_lower == "title":
                items[key]["title"] = value or ""
            elif field_lower == "doi":
                items[key]["DOI"] = value or ""
            elif field_lower == "url":
                items[key]["url"] = value or ""

    except Exception:
        pass

    conn.close()
    return items


# -----------------------------
# Zotero field parsing
# -----------------------------

def extract_json_from_field_code(code):
    start = code.find("{")
    if start == -1:
        return None

    depth = 0
    in_string = False
    escape = False

    for i in range(start, len(code)):
        ch = code[i]

        if escape:
            escape = False
            continue

        if ch == "\\":
            escape = True
            continue

        if ch == '"':
            in_string = not in_string
            continue

        if in_string:
            continue

        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return code[start:i + 1]

    return None


def parse_citation_json(code):
    raw = extract_json_from_field_code(code)

    if not raw:
        return None

    try:
        return json.loads(raw)
    except Exception:
        return None


def get_plain_citation_from_field_code(code):
    data = parse_citation_json(code)

    if data:
        props = data.get("properties", {}) or {}
        plain = props.get("plainCitation") or props.get("formattedCitation")
        if plain:
            return plain

    match = re.search(r'"plainCitation"\s*:\s*"([^"]+)"', code)
    if match:
        return match.group(1)

    match = re.search(r'"formattedCitation"\s*:\s*"([^"]+)"', code)
    if match:
        return match.group(1)

    return ""


def citation_item_label(item_data, key=None):
    title = item_data.get("title", "")

    year = ""
    issued = item_data.get("issued", {})
    try:
        year = str(issued["date-parts"][0][0])
    except Exception:
        pass

    authors = item_data.get("author", [])
    author_label = ""

    if authors:
        first = authors[0]
        family = first.get("family") or first.get("literal") or ""
        if len(authors) == 1:
            author_label = family
        else:
            author_label = f"{family} et al."

    bits = []
    if author_label:
        bits.append(author_label)
    if year:
        bits.append(year)

    if bits:
        label = ", ".join(bits)
    elif title:
        label = title
    elif key:
        label = f"Zotero item key {key}"
    else:
        label = "Unknown citation item"

    if title and title.lower() not in label.lower():
        label = f"{label} — {short_text(title, 80)}"

    return short_text(label, 150)


def parse_citation_items(code):
    data = parse_citation_json(code)
    parsed = []

    if data:
        citation_items = data.get("citationItems", [])

        for item in citation_items:
            key = None

            uris = item.get("uris", [])
            for uri in uris:
                match = re.search(r"/items/([A-Z0-9]+)", uri)
                if match:
                    key = match.group(1)
                    break

            item_data = item.get("itemData", {}) or {}

            if not key:
                key = item.get("itemKey") or item_data.get("key")

            parsed.append({
                "key": key,
                "label": citation_item_label(item_data, key),
                "title": item_data.get("title", ""),
            })

    if not parsed:
        for key in re.findall(r"/items/([A-Z0-9]+)", code):
            parsed.append({
                "key": key,
                "label": f"Zotero item key {key}",
            })

    return parsed


# -----------------------------
# Citation-like text detection
# -----------------------------

def parenthetical_author_year_pattern():
    return re.compile(
        r"\("
        r"(?=[^()\r\n]{1,220}\b(?:19|20)\d{2}[a-z]?\b)"
        r"(?=[^()\r\n]{1,220}(?:\bet\s+al\.?\b|[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s*,))"
        r"[^()\r\n]{1,220}"
        r"\)"
    )


def narrative_et_al_year_pattern():
    return re.compile(
        r"(?<!\w)"
        r"("
        r"(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)"
        r"(?:[\s,]+(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)){0,5}"
        r"\s+et\s+al\.?"
        r")"
        r"\s*\("
        r"((?:19|20)\d{2}[a-z]?)"
        r"\)"
    )


def narrative_two_author_year_pattern():
    return re.compile(
        r"(?<!\w)"
        r"("
        r"(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)"
        r"(?:[\s,]+(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)){0,3}"
        r"\s+(?:and|&)\s+"
        r"(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)"
        r"(?:[\s,]+(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)){0,3}"
        r")"
        r"\s*\("
        r"((?:19|20)\d{2}[a-z]?)"
        r"\)"
    )


def narrative_single_author_year_pattern():
    return re.compile(
        r"(?<!\w)"
        r"("
        r"(?:[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+|[A-Z]\.)"
        r")"
        r"\s*\("
        r"((?:19|20)\d{2}[a-z]?)"
        r"\)"
    )


def all_manual_detection_patterns():
    return [
        parenthetical_author_year_pattern(),
        narrative_et_al_year_pattern(),
        narrative_two_author_year_pattern(),
        narrative_single_author_year_pattern(),
    ]


def split_author_year_citation(citation_text):
    text = citation_text.strip()

    narrative_match = re.match(
        r"^(.*?)\s*\(((?:19|20)\d{2}[a-z]?)\)$",
        text
    )

    if narrative_match and not text.startswith("("):
        author = narrative_match.group(1).strip(" ,")
        year = narrative_match.group(2)
        raw = f"{author} ({year})"
        norm = normalize_citation_component(f"{author}, {year}")
        return [{"raw": raw, "norm": norm}] if norm else []

    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]

    parts = [p.strip() for p in text.split(";") if p.strip()]
    results = []

    for part in parts:
        years = re.findall(r"(?:19|20)\d{2}[a-z]?", part)
        if not years:
            continue

        first_year_match = re.search(r"(?:19|20)\d{2}[a-z]?", part)
        author = part[:first_year_match.start()].strip(" ,") if first_year_match else ""

        for year in years:
            raw = f"{author}, {year}".strip(" ,")
            norm = normalize_citation_component(raw)
            if norm:
                results.append({
                    "raw": raw,
                    "norm": norm,
                })

    return results


def should_ignore_apparent_citation(text):
    stripped = text.strip()
    inside = stripped[1:-1].strip().lower() if stripped.startswith("(") and stripped.endswith(")") else stripped.lower()

    # Ignore any year/date range. Manual citation checking only accepts a single year.
    if re.search(r"(?:19|20)\d{2}\s*(?:[-–—/]|to)\s*(?:(?:19|20)\d{2}|\d{2})", inside):
        return True

    # Ignore single year only, e.g. (2024).
    if re.fullmatch(r"(?:19|20)\d{2}[a-z]?", inside):
        return True

    excluded = (
        "fig.",
        "figure",
        "table",
        "section",
        "chapter",
        "eq.",
        "equation",
        "box",
        "annex",
        "supplement",
    )

    if inside.startswith(excluded):
        return True

    if len(text) > 300:
        return True

    has_author_signal = (
        re.search(r"\bet\s+al\.?\b", inside)
        or re.search(r"\b[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s*,\s*(?:19|20)\d{2}[a-z]?", inside)
        or re.search(r"\b[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s+(?:and|&)\s+[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+", inside)
        or re.search(r"\b[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s*\((?:19|20)\d{2}[a-z]?\)", inside)
    )

    if not has_author_signal:
        return True

    return False


# -----------------------------
# Word helpers
# -----------------------------

def copy_to_safe_temp(source_path):
    source_path = clean_path(source_path)

    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Input document not found:\n{source_path}")

    temp_dir = tempfile.mkdtemp(prefix="zotero_refcheck_")
    temp_path = os.path.join(temp_dir, os.path.basename(source_path))
    shutil.copy2(source_path, temp_path)
    return temp_path, temp_dir


def find_bibliography_field(doc):
    for field in doc.Fields:
        code = field.Code.Text
        if "ADDIN ZOTERO_BIBL" in code:
            return field
    return None


def add_comment(doc, word_range, text):
    try:
        doc.Comments.Add(Range=word_range, Text=text)
    except Exception:
        pass


def open_word_document(word, path):
    path = clean_path(path)

    if not os.path.exists(path):
        raise FileNotFoundError(f"Word file not found:\n{path}")

    try:
        return word.Documents.Open(
            FileName=path,
            ReadOnly=False,
            AddToRecentFiles=False,
            ConfirmConversions=False,
            Visible=False,
        )
    except Exception as e:
        raise RuntimeError(
            "Microsoft Word could not open this document cleanly.\n\n"
            "Please open the original file manually in Word, save it once, close Word, "
            "then run this checker again.\n\n"
            f"File:\n{path}\n\n"
            f"Word error:\n{e}"
        )


def close_word_safely(word, doc=None):
    if doc is not None:
        try:
            doc.Close(False)
        except Exception:
            pass

    if word is not None:
        try:
            word.Quit()
        except Exception:
            pass


def build_zotero_result_ranges(doc):
    ranges = []

    for field in doc.Fields:
        try:
            code = field.Code.Text
            code_upper = code.upper()

            # Include valid and damaged Zotero fields so their visible text is
            # not later misclassified as a manually typed citation.
            if "ADDIN ZOTERO" not in code_upper:
                continue

            ranges.append((field.Result.Start, field.Result.End))

        except Exception:
            pass

    ranges.sort(key=lambda x: x[0])
    return ranges

def candidate_overlaps_zotero_range(candidate_range, zotero_ranges):
    try:
        start = candidate_range.Start
        end = candidate_range.End
    except Exception:
        return False

    # Ranges are sorted. Break once we have passed this candidate.
    for z_start, z_end in zotero_ranges:
        if z_start >= end:
            break
        if ranges_overlap(start, end, z_start, z_end):
            return True
    return False


def find_visible_range_in_paragraph(paragraph_range, visible_text, preferred_start=None):
    try:
        search_range = paragraph_range.Duplicate

        if preferred_start is not None:
            # Start near the regex match, but rely on Word Find for the actual range.
            search_range.Start = max(paragraph_range.Start, paragraph_range.Start + preferred_start - 5)

        find = search_range.Find
        find.ClearFormatting()
        find.Text = visible_text
        find.MatchCase = False
        find.MatchWholeWord = False

        if find.Execute():
            return search_range
    except Exception:
        pass

    # Fallback to paragraph range; this is less precise but still attaches a useful comment.
    return paragraph_range


def find_component_range_in_paragraph(paragraph_range, raw_text, full_match_text=None, preferred_start=None):
    # First try exact component, e.g. Jenkins et al. (2022) or Jenkins et al., 2022.
    r = find_visible_range_in_paragraph(paragraph_range, raw_text, preferred_start)
    try:
        if short_text(r.Text, len(raw_text) + 5).lower().find(raw_text.lower()) != -1:
            return r
    except Exception:
        pass

    # Then try the full parenthetical/narrative match if provided.
    if full_match_text:
        return find_visible_range_in_paragraph(paragraph_range, full_match_text, preferred_start)

    return r


# -----------------------------
# Manual citation detection, optimized
# -----------------------------

def check_manual_citations_in_main_doc(doc, zotero_ranges, progress_callback=None):
    if progress_callback:
        progress_callback("Checking possible manually typed in-text citations...")

    patterns = all_manual_detection_patterns()

    for paragraph in doc.Paragraphs:
        try:
            paragraph_text = paragraph.Range.Text
            paragraph_range = paragraph.Range
        except Exception:
            continue

        if not any(p.search(paragraph_text) for p in patterns):
            continue

        manual_candidates = []

        for pattern in patterns:
            for match in pattern.finditer(paragraph_text):
                citation_text = match.group(0)

                if should_ignore_apparent_citation(citation_text):
                    continue

                candidate_range = find_visible_range_in_paragraph(
                    paragraph_range,
                    citation_text,
                    preferred_start=match.start()
                )

                # If the visible citation occurrence is backed by a Zotero field, it is linked, not manual.
                if candidate_overlaps_zotero_range(candidate_range, zotero_ranges):
                    continue

                comps = split_author_year_citation(citation_text)

                for comp in comps:
                    manual_candidates.append({
                        "raw": comp["raw"],
                        "norm": comp["norm"],
                        "match_start": match.start(),
                        "match_end": match.end(),
                        "full_match": citation_text,
                    })

        if not manual_candidates:
            continue

        seen = set()

        for item in manual_candidates:
            unique_id = (item["norm"], item["match_start"], item["match_end"])

            if unique_id in seen:
                continue

            seen.add(unique_id)

            target_range = find_component_range_in_paragraph(
                paragraph_range,
                item["raw"],
                full_match_text=item["full_match"],
                preferred_start=item["match_start"]
            )

            log_comment_context("Possible manually typed in-text citation", {
                "possible_manual_citation": item["raw"],
                "paragraph_excerpt": short_text(paragraph_text, 180),
            })

            add_comment(
                doc,
                target_range,
                "Issue: This appears to be a manually typed in-text citation. "
                "Please insert this citation using Zotero and update the reference list.\n\n"
                f"Possible manual citation: {item['raw']}\n\n"
                f"Paragraph excerpt: {short_text(paragraph_text, 180)}"
            )


# -----------------------------
# Main processing
# -----------------------------

def process_document(input_path, checked_output_path, db_path, selected_library, progress_callback=None):
    def progress(message):
        if progress_callback:
            progress_callback(message)

    global debug_path

    temp_dir = None

    # Save the debug log in the same folder as the selected Word document.
    debug_path = os.path.join(
        os.path.dirname(clean_path(input_path)),
        "refcheck_debug.txt"
    )

    start_debug_run(input_path, selected_library)
    total_started_at = time.perf_counter()

    progress("Loading Zotero database...")
    items = load_zotero_items(db_path)
    allowed_libraries = {DEFAULT_LIBRARY, selected_library}

    items = {
        key: meta
        for key, meta in items.items()
        if meta.get("library") in allowed_libraries
    }

    # Build this once so a citation whose embedded Zotero key is no longer
    # present can be resolved quickly by its normalized embedded title.
    title_lookup = {}

    for candidate_key, meta in items.items():
        candidate_title = normalize_match_text(meta.get("title", ""))

        if candidate_title:
            title_lookup.setdefault(candidate_title, []).append(candidate_key)

    progress("Copying document to a safe local working folder...")
    temp_doc, temp_dir = copy_to_safe_temp(input_path)
    temp_checked_output = os.path.join(temp_dir, "checked_output.docx")

    progress("Starting Microsoft Word...")
    word = None
    doc = None
    used_keys = set()
    used_titles = set()

    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        progress("Opening Word document...")
        doc = open_word_document(word, temp_doc)

        progress("Checking linked Zotero in-text citations...")

        for field in doc.Fields:
            code = field.Code.Text
            code_upper = code.upper()

            # Ignore non-Zotero fields.
            if "ADDIN ZOTERO" not in code_upper:
                continue

            # The bibliography is a valid Zotero field, but it is not an
            # in-text citation and should not be reported as damaged.
            if "ZOTERO_BIBL" in code_upper:
                continue

            # Any other Zotero field that is not a valid ZOTERO_ITEM field is
            # treated as damaged or incomplete, for example "ADDIN Zotero TEMPP".
            if "ZOTERO_ITEM" not in code_upper:
                visible_text = range_text(field.Result)

                log_comment_context(
                    "Damaged Zotero field",
                    {
                        "field_code": short_text(code, 120),
                        "visible_text": visible_text,
                    },
                )

                add_comment(
                    doc,
                    field.Result,
                    "Issue: This appears to be a damaged Zotero citation field. "
                    "The citation may have become corrupted or partially unlinked. "
                    "Please reinsert the citation using Zotero."
                )
                continue

            plain = get_plain_citation_from_field_code(code) or range_text(field.Result)
            citation_items = parse_citation_items(code)

            if not citation_items:
                log_comment_context("Unreadable linked Zotero citation", {
                    "citation": plain,
                })
                add_comment(
                    doc,
                    field.Result,
                    f"Issue: This is a linked Zotero citation, but the checker could not read its Zotero item data. "
                    f"It may be a damaged or partially copied Zotero field.\n\nCitation: {plain}"
                )
                continue

            for citation_item in citation_items:
                key = citation_item.get("key")
                label = citation_item.get("label", "Unknown citation item")

                if not key:
                    log_comment_context("Missing item key in linked Zotero citation", {
                        "citation": plain,
                        "citation_item": label,
                    })
                    add_comment(
                        doc,
                        field.Result,
                        f"Issue: This is a linked Zotero citation, but the checker could not read the item key for one citation item. "
                        f"It may be a damaged or partially copied Zotero field.\n\nCitation item: {label}\n\nCitation: {plain}"
                    )
                    continue

                used_keys.add(key)

                embedded_title = normalize_match_text(
                    citation_item.get("title", "")
                )
                resolved_key = key

                # If the exact embedded key is absent, look for the same item
                # title in the two selected libraries. This handles copied,
                # merged, recreated, or duplicated Zotero items whose keys differ.
                if key not in items and embedded_title:
                    alternate_keys = title_lookup.get(embedded_title, [])

                    if alternate_keys:
                        resolved_key = alternate_keys[0]
                        used_keys.add(resolved_key)

                title = ""

                if resolved_key in items:
                    title = normalize_match_text(
                        items[resolved_key].get("title", "")
                    )

                if not title:
                    title = embedded_title

                if title:
                    used_titles.add(title)

                if resolved_key not in items:
                    log_comment_context("Linked Zotero citation item not found", {
                        "citation": plain,
                        "citation_item": label,
                        "embedded_key": key,
                        "embedded_title": citation_item.get("title", ""),
                        "key_in_filtered_items": key in items,
                        "title_match_in_selected_libraries": False,
                    })
                    add_comment(
                        doc,
                        field.Result,
                        f"Issue: This linked Zotero citation item was not found in the selected local Zotero libraries. "
                        f"The link may be broken, the item may have been deleted, or the relevant group library may not be synced.\n\n"
                        f"Citation item: {label}\n\nCitation: {plain}"
                    )
                    continue

                library = items[resolved_key]["library"]

                if library not in allowed_libraries:
                    log_comment_context("Unexpected Zotero library", {
                        "citation": plain,
                        "citation_item": label,
                        "embedded_key": key,
                        "resolved_key": resolved_key,
                        "library": library,
                        "allowed_libraries": sorted(allowed_libraries),
                    })
                    add_comment(
                        doc,
                        field.Result,
                        f"Issue: Citation item comes from unexpected Zotero library: {library}.\n\nCitation item: {label}"
                    )


        progress("Checking Zotero bibliography...")
        bibliography_field = find_bibliography_field(doc)

        if bibliography_field:
            bib_range = bibliography_field.Result
            paragraphs = bib_range.Paragraphs

            # Duplicate checks use the normalized full bibliography entry as a
            # safeguard against false positives from parent book/report titles.
            seen_bibliography_keys = {}
            seen_bibliography_titles = {}

            for paragraph in paragraphs:

                anchor = paragraph_anchor_range(paragraph)

                try:
                    entry_text = anchor.Text
                except Exception:
                    entry_text = ""

                p_text = normalize_match_text(entry_text)
                


                if not p_text:
                    continue

                matched_key = None
                matched_score = 0

                for key, meta in items.items():
                    title = normalize_match_text(meta.get("title", ""))

                    if not title or len(title) <= 25 or title not in p_text:
                        continue

                    # Prefer titles that appear earlier in the bibliography entry.
                    # This helps with chapters/annexes where the entry contains both
                    # the chapter title and the parent book/report title after "In ...".
                    position = p_text.find(title)
                    score = len(title) - position

                    if score > matched_score:
                        matched_key = key
                        matched_score = score

                if not matched_key:
                    continue

                meta = items[matched_key]
                entry_preview = short_text(entry_text, 100)
                matched_title = normalize_match_text(meta.get("title", ""))

                previous_key_entry = seen_bibliography_keys.get(matched_key)
                if previous_key_entry and previous_key_entry["entry_text"] == p_text:
                    log_comment_context("Duplicate bibliography entry: same Zotero item", {
                        "entry": entry_preview,
                        "matched_key": matched_key,
                        "matched_title": matched_title,
                        "first_entry": previous_key_entry["entry_preview"],
                    })
                    add_comment(
                        doc,
                        anchor,
                        "Issue: This bibliography entry duplicates another entry in the "
                        "reference list and matches the same Zotero item.\n\n"
                        f"Entry: {entry_preview}"
                    )
                else:
                    seen_bibliography_keys.setdefault(
                        matched_key,
                        {
                            "entry_text": p_text,
                            "entry_preview": entry_preview,
                        },
                    )

                previous_title_entry = seen_bibliography_titles.get(matched_title)
                if (
                    matched_title
                    and previous_title_entry
                    and previous_title_entry["key"] != matched_key
                    and previous_title_entry["entry_text"] == p_text
                ):
                    log_comment_context(
                        "Duplicate bibliography entry: same title, different Zotero items",
                        {
                            "entry": entry_preview,
                            "matched_title": matched_title,
                            "current_key": matched_key,
                            "first_key": previous_title_entry["key"],
                            "first_entry": previous_title_entry["entry_preview"],
                        },
                    )
                    add_comment(
                        doc,
                        anchor,
                        "Issue: This bibliography entry has the same title as another "
                        "entry but matches a different Zotero item. This may indicate "
                        "duplicate Zotero records.\n\n"
                        f"Entry: {entry_preview}"
                    )
                elif matched_title:
                    seen_bibliography_titles.setdefault(
                        matched_title,
                        {
                            "key": matched_key,
                            "entry_text": p_text,
                            "entry_preview": entry_preview,
                        },
                    )

                # Safety check: even if matched_key picked the parent book/report title,
                # do not mark the bibliography entry as uncited if the entry contains
                # any title that was found in linked in-text Zotero citations.
                entry_contains_used_title = any(
                    used_title
                    and (
                        len(used_title) > 15
                        or used_title in {"glossary"}
                    )
                    and used_title in p_text
                    for used_title in used_titles
                )

                if (
                    matched_key not in used_keys
                    and matched_title not in used_titles
                    and not entry_contains_used_title
                ):
                    log_comment_context("Bibliography entry not found as linked in-text citation", {
                        "entry": entry_preview,
                        "matched_key": matched_key,
                        "matched_title": matched_title,
                        "key_in_used_keys": matched_key in used_keys,
                        "title_in_used_titles": matched_title in used_titles,
                        "entry_contains_used_title": entry_contains_used_title,
                    })

                    add_comment(
                        doc,
                        anchor,
                        f"Issue: This bibliography entry appears in the reference list but was not found as a linked in-text citation.\n\nEntry: {entry_preview}"
                    )



                #if not meta.get("DOI") and not meta.get("url"):
                #    add_comment(
                #        doc,
                #        anchor,
                #        f"Issue: This reference has neither DOI nor URL in Zotero.\n\nEntry: {entry_preview}"
                #    )

                entry_has_doi = bool(
                    re.search(r"\b10\.\d{4,9}/\S+", entry_text, re.I)
                )

                entry_has_url = (
                    "http://" in entry_text.lower()
                    or "https://" in entry_text.lower()
                )

                if (
                    not meta.get("DOI")
                    and not meta.get("url")
                    and not entry_has_doi
                    and not entry_has_url
                ):
                    log_comment_context("Reference has neither DOI nor URL", {
                        "entry": entry_preview,
                        "matched_key": matched_key,
                        "matched_title": matched_title,
                        "zotero_doi_field": meta.get("DOI", ""),
                        "zotero_url_field": meta.get("url", ""),
                        "entry_has_doi_text": entry_has_doi,
                        "entry_has_url_text": entry_has_url,
                    })

                    add_comment(
                        doc,
                        anchor,
                        f"Issue: This reference has neither DOI nor URL in Zotero.\n\nEntry: {entry_preview}"
                    )


        else:
            log_comment_context("Zotero bibliography field not found", {
                "document": input_path,
            })
            add_comment(
                doc,
                doc.Content,
                "Issue: Zotero bibliography field was not found. The bibliography may be unlinked or manually edited."
            )


        progress("Indexing linked Zotero citation ranges...")
        zotero_ranges = build_zotero_result_ranges(doc)

        check_manual_citations_in_main_doc(doc, zotero_ranges, progress_callback=progress)

        progress("Saving checked Word document...")
        doc.SaveAs2(
            FileName=clean_path(temp_checked_output),
            FileFormat=WD_FORMAT_XML_DOCUMENT,
            AddToRecentFiles=False,
        )

        close_word_safely(word, doc)
        word = None
        doc = None

        shutil.copy2(temp_checked_output, clean_path(checked_output_path))
        log_elapsed_time("Total reference check", total_started_at)

    finally:
        close_word_safely(word, doc)

        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass

    progress("Finished.")


# -----------------------------
# GUI
# -----------------------------

def main():
    root = tk.Tk()
    root.title("Zotero Reference Checker")
    root.geometry("620x260")
    root.resizable(False, False)

    db_path = find_zotero_db()

    if not db_path:
        messagebox.showwarning(
            "Zotero database not found",
            "Could not automatically find zotero.sqlite. Please select it manually.",
        )
        db_path = choose_zotero_db()

    if not db_path:
        messagebox.showerror("Error", "No Zotero database selected.")
        root.destroy()
        return

    libraries = sorted(set(load_library_names(db_path).values()))

    if DEFAULT_LIBRARY in libraries:
        libraries.remove(DEFAULT_LIBRARY)

    selected_library = tk.StringVar()

    setup_frame = tk.Frame(root)
    setup_frame.pack(fill="both", expand=True, padx=20, pady=20)

    tk.Label(
        setup_frame,
        text=f"Default library always included:\n{DEFAULT_LIBRARY}",
        justify="center",
    ).pack(pady=(0, 10))

    tk.Label(
        setup_frame,
        text="Select second Zotero group library:",
    ).pack(pady=(0, 5))

    dropdown = ttk.Combobox(
        setup_frame,
        textvariable=selected_library,
        width=70,
        state="readonly",
    )
    dropdown["values"] = libraries
    dropdown.pack(pady=(0, 10))

    if libraries:
        dropdown.current(0)

    progress_frame = tk.Frame(root)

    progress_label = tk.Label(
        progress_frame,
        text="",
        wraplength=560,
        justify="center",
        height=3,
    )
    progress_label.pack(pady=(25, 15))

    progress_bar = ttk.Progressbar(
        progress_frame,
        mode="indeterminate",
        length=480,
    )
    progress_bar.pack(pady=(0, 15))

    def set_progress(text):
        progress_label.config(text=text)
        root.update_idletasks()
        root.update()

    def run_check():
        if not selected_library.get():
            messagebox.showerror("Error", "Please select the second Zotero library.")
            return

        input_path = filedialog.askopenfilename(
            title="Select Word document",
            filetypes=[("Word documents", "*.docx")],
        )

        if not input_path:
            return

        setup_frame.pack_forget()
        progress_frame.pack(fill="both", expand=True, padx=20, pady=20)

        progress_label.config(text="Starting reference check...")
        progress_bar.start(15)
        root.update_idletasks()

        input_path_clean = clean_path(input_path)
        base, ext = os.path.splitext(input_path_clean)
        checked_output_path = base + "_checked" + ext

        try:
            process_document(
                input_path=input_path_clean,
                checked_output_path=checked_output_path,
                db_path=db_path,
                selected_library=selected_library.get(),
                progress_callback=set_progress,
            )

            progress_bar.stop()
            progress_label.config(text="Finished. Output document has been created.")
            root.update_idletasks()

            messagebox.showinfo(
                "Reference check complete",
                f"Finished.\n\nChecked file:\n{checked_output_path}",
            )

            root.destroy()

        except Exception as e:
            progress_bar.stop()
            messagebox.showerror("Error", str(e))

            progress_frame.pack_forget()
            setup_frame.pack(fill="both", expand=True, padx=20, pady=20)

    tk.Button(
        setup_frame,
        text="Select Word Document and Run Check",
        command=run_check,
        width=35,
    ).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
