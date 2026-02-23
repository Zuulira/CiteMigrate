#!/usr/bin/env python3
"""
CiteMigrate - macOS App (PyQt6)
Converts Citavi citation field codes in Word documents (.docx) to
Zotero-compatible field codes with proper linkage to your Zotero library.

"Citavi" is a trademark of Swiss Academic Software GmbH.
"Zotero" is a registered trademark of the Corporation for Digital Scholarship.
"Microsoft Word" is a trademark of Microsoft Corporation.
This tool is not affiliated with, endorsed by, or produced by any of these
organizations. Trademark names are used solely for descriptive purposes.

License: GPLv3 (see LICENSE file)

Usage:
    python citemigrate.py

Or build as .app:
    bash build_app.sh
"""

import atexit
import base64
import binascii
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import uuid
import zipfile

# ──────────────────────────────────────────────
# GUI imports (PyQt6 - no Tcl/Tk dependency)
# ──────────────────────────────────────────────
try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
        QMessageBox, QProgressBar, QCheckBox, QComboBox, QGroupBox,
        QFrame, QListWidget, QListWidgetItem,
        QAbstractItemView, QTabWidget, QScrollArea
    )
    from PyQt6.QtCore import Qt, QThread, pyqtSignal
    from PyQt6.QtGui import QFont, QIcon
except ImportError:
    sys.exit(
        "ERROR: PyQt6 is required.\n"
        "Install with: pip install PyQt6\n"
        "Or if using the build script: bash build_app.sh"
    )

try:
    from lxml import etree
except ImportError:
    sys.exit("ERROR: lxml is required. Install with: pip install lxml")

try:
    from pyzotero import zotero
except ImportError:
    sys.exit("ERROR: pyzotero is required. Install with: pip install pyzotero")


# ══════════════════════════════════════════════
# CITAVI PARSING & ZOTERO CONVERSION ENGINE
# ══════════════════════════════════════════════

DEFAULT_STYLE_URI = "http://www.zotero.org/styles/harvard-cite-them-right"

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

for _prefix, _uri in NAMESPACES.items():
    etree.register_namespace(_prefix, _uri)


# ── Citavi Parsing ────────────────────────────

def find_citavi_sdts(tree):
    sdts = []
    for sdt in tree.iter(f"{{{NAMESPACES['w']}}}sdt"):
        sdt_pr = sdt.find(f"{{{NAMESPACES['w']}}}sdtPr")
        if sdt_pr is None:
            continue
        tag_elem = sdt_pr.find(f"{{{NAMESPACES['w']}}}tag")
        if tag_elem is None:
            continue
        tag_val = tag_elem.get(f"{{{NAMESPACES['w']}}}val", "")
        if tag_val.startswith("CitaviPlaceholder#"):
            sdts.append((sdt, tag_val))
    return sdts


def decode_citavi_payload(sdt):
    sdt_content = sdt.find(f"{{{NAMESPACES['w']}}}sdtContent")
    if sdt_content is None:
        return None

    instr_parts = []
    for instr in sdt_content.iter(f"{{{NAMESPACES['w']}}}instrText"):
        if instr.text:
            instr_parts.append(instr.text.strip())

    if not instr_parts:
        return None

    raw = "".join(instr_parts)

    for candidate in [raw, raw.split(None, 1)[-1] if " " in raw else raw]:
        candidate = candidate.strip()
        if not candidate:
            continue
        try:
            decoded = base64.b64decode(candidate).decode("utf-8", errors="replace")
            return json.loads(decoded)
        except (json.JSONDecodeError, ValueError, UnicodeDecodeError, binascii.Error):
            pass

    json_match = re.search(r'\{.*\}', raw, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass

    return None


def extract_citavi_display_text(sdt):
    sdt_content = sdt.find(f"{{{NAMESPACES['w']}}}sdtContent")
    if sdt_content is None:
        return ""

    texts = []
    in_display = False
    for elem in sdt_content.iter():
        tag = etree.QName(elem.tag).localname if isinstance(elem.tag, str) else ""
        if tag == "fldChar":
            fld_type = elem.get(f"{{{NAMESPACES['w']}}}fldCharType", "")
            if fld_type == "separate":
                in_display = True
            elif fld_type == "end":
                in_display = False
        if in_display and tag == "t" and elem.text:
            texts.append(elem.text)

    return "".join(texts).strip()


def extract_citation_info_from_citavi(payload):
    citations = []
    if payload is None:
        return citations

    def extract_from_entry(entry):
        info = {}
        for key in ["Title", "title", "TitleString"]:
            if key in entry and entry[key]:
                info["title"] = str(entry[key]).strip()
                break

        authors = []
        for key in ["Authors", "authors", "AuthorsOrEditorsOrOrganizations"]:
            if key in entry and entry[key]:
                author_list = entry[key]
                if isinstance(author_list, list):
                    for a in author_list:
                        if isinstance(a, dict):
                            family = a.get("LastName", a.get("Family", a.get("family", "")))
                            given = a.get("FirstName", a.get("Given", a.get("given", "")))
                            authors.append({"family": family, "given": given})
                        elif isinstance(a, str):
                            parts = a.split(",")
                            if len(parts) >= 2:
                                authors.append({"family": parts[0].strip(), "given": parts[1].strip()})
                            else:
                                authors.append({"family": a.strip(), "given": ""})
                break

        if authors:
            info["authors"] = authors

        for key in ["Year", "year", "YearResolved", "Date"]:
            if key in entry and entry[key]:
                year_match = re.search(r'\b(1[0-9]{3}|2[0-9]{3})\b', str(entry[key]))
                if year_match:
                    info["year"] = year_match.group(1)
                break

        for key in ["Doi", "doi", "DOI"]:
            if key in entry and entry[key]:
                info["doi"] = str(entry[key]).strip()
                break

        for key in ["Isbn", "isbn", "ISBN"]:
            if key in entry and entry[key]:
                info["isbn"] = str(entry[key]).strip()
                break

        return info if info else None

    entries = []
    if isinstance(payload, dict):
        for key in ["Entries", "entries", "Citations", "citations", "References"]:
            if key in payload and isinstance(payload[key], list):
                entries = payload[key]
                break

        if not entries:
            result = extract_from_entry(payload)
            if result:
                return [result]
            for val in payload.values():
                if isinstance(val, list):
                    for item in val:
                        if isinstance(item, dict):
                            result = extract_from_entry(item)
                            if result:
                                citations.append(result)
                    if citations:
                        return citations
    elif isinstance(payload, list):
        entries = payload

    for entry in entries:
        if isinstance(entry, dict):
            result = extract_from_entry(entry)
            if result:
                citations.append(result)

    return citations


# ── Zotero Matching ───────────────────────────

class ZoteroMatcher:
    def __init__(self, library_id, api_key, library_type="user"):
        self.zot = zotero.Zotero(library_id, library_type, api_key)
        self._cache = {}
        self._all_items = None

    def _get_all_items(self):
        if self._all_items is not None:
            return self._all_items

        self._all_items = []
        start = 0
        limit = 100
        while True:
            # Use top() to get only top-level items (excludes attachments & notes)
            # This avoids URL-encoding issues with itemType filter parameters
            items = self.zot.top(start=start, limit=limit)
            if not items:
                break
            self._all_items.extend(items)
            start += limit
            if len(items) < limit:
                break

        return self._all_items

    def get_item_count(self):
        return len(self._get_all_items())

    def find_match(self, citation_info):
        cache_key = json.dumps(citation_info, sort_keys=True)
        if cache_key in self._cache:
            return self._cache[cache_key]

        items = self._get_all_items()

        if "doi" in citation_info:
            doi = citation_info["doi"].lower().strip()
            for item in items:
                item_doi = item["data"].get("DOI", "").lower().strip()
                if item_doi and item_doi == doi:
                    self._cache[cache_key] = item
                    return item

        if "isbn" in citation_info:
            isbn = re.sub(r'[\s-]', '', citation_info["isbn"])
            for item in items:
                item_isbn = re.sub(r'[\s-]', '', item["data"].get("ISBN", ""))
                if item_isbn and item_isbn == isbn:
                    self._cache[cache_key] = item
                    return item

        best_match = None
        best_score = 0
        target_title = (citation_info.get("title", "") or "").lower()
        target_year = citation_info.get("year", "")
        target_authors = citation_info.get("authors", [])
        target_family = target_authors[0].get("family", "").lower() if target_authors else ""

        for item in items:
            score = 0
            data = item["data"]
            item_date = data.get("date", "")
            if target_year and target_year in str(item_date):
                score += 2

            item_creators = data.get("creators", [])
            if target_family and item_creators:
                for creator in item_creators:
                    creator_last = creator.get("lastName", "").lower()
                    # Also check "name" field (used for organization/institutional authors)
                    creator_name = creator.get("name", "").lower()
                    if (creator_last and creator_last == target_family) or \
                       (creator_name and (creator_name == target_family or
                        target_family in creator_name)):
                        score += 3
                        break

            item_title = data.get("title", "").lower()
            if target_title and item_title:
                target_words = set(target_title.split())
                item_words = set(item_title.split())
                if target_words and item_words:
                    overlap = len(target_words & item_words)
                    ratio = overlap / max(len(target_words), len(item_words))
                    if ratio > 0.6:
                        score += 5 * ratio

            if score > best_score and score >= 3:
                best_score = score
                best_match = item

        self._cache[cache_key] = best_match
        return best_match

    def find_match_by_display_text(self, display_text):
        if not display_text:
            return []

        text = display_text.strip("()[] ")
        parts = [p.strip() for p in text.split(";")]
        matches = []
        for part in parts:
            part = part.strip()
            if not part:
                continue

            # Try to extract year (4-digit number) from the citation part
            year_match = re.search(r'(\d{4})', part)
            if not year_match:
                continue
            year = year_match.group(1)

            # Everything before the year is the author portion
            author_portion = part[:year_match.start()].strip().rstrip(",").strip()
            # Remove "et al." / "et al" from author portion
            author_portion = re.sub(r'\s*et\s+al\.?\s*$', '', author_portion, flags=re.IGNORECASE).strip()

            if not author_portion:
                continue

            # Strip common citation prefixes in multiple languages:
            # German: "vgl." (vergleiche), "s." (siehe), "s.a." (siehe auch)
            # English: "cf.", "see", "see also"
            # French: "cf.", "voir", "voir aussi"
            # Spanish: "véase", "cf."
            # Italian: "cfr.", "vedi"
            # Portuguese: "cf.", "ver", "veja"
            # Dutch: "vgl.", "zie"
            author_portion = re.sub(
                r'^(?:vgl\.|cf\.|cfr\.|see(?:\s+also)?|voir(?:\s+aussi)?|'
                r's\.(?:a\.)?|véase|vedi|ver\b|veja|zie)\s*',
                '', author_portion, flags=re.IGNORECASE
            ).strip()

            # Extract primary author by splitting on "&", " and ", ","
            # For " und " (German): only split if it separates short name-like parts
            # (to avoid breaking org names like "National Institute of Health und Medical Research")
            # Also handle French "et", Spanish "y", Italian "e", Dutch "en"
            primary_author = re.split(
                r'\s*&\s*|\s+and\s+|\s*,\s*',
                author_portion, flags=re.IGNORECASE
            )[0].strip()
            # Now try splitting on connectors only if both sides look like short author names
            for connector in [r'\s+und\s+', r'\s+et\s+', r'\s+y\s+', r'\s+e\s+', r'\s+en\s+']:
                parts = re.split(connector, primary_author, flags=re.IGNORECASE)
                if len(parts) > 1 and len(parts[0].split()) <= 2:
                    primary_author = parts[0].strip()
                    break

            if not primary_author:
                continue

            # Handle names with trailing initials like "Miller J." or "John Smith"
            # Strip trailing single-letter initials (with or without period)
            primary_author = re.sub(r'\s+[A-Z]\.?\s*$', '', primary_author).strip()

            # If primary_author has multiple words (e.g. "John Smith" or
            # "World Health Organization"), try matching with:
            # 1. The full string as family name (for organizations)
            # 2. The last word as family name (for "FirstName LastName" patterns)
            candidates = [primary_author]
            words = primary_author.split()
            if len(words) > 1:
                candidates.append(words[-1])  # last word as family name

            result = None
            for candidate in candidates:
                citation_info = {
                    "authors": [{"family": candidate, "given": ""}],
                    "year": year
                }
                result = self.find_match(citation_info)
                if result:
                    break

            # If still no match and we have a multi-word name, try as organization
            # or title-based search
            if not result and len(words) > 2:
                # Try the full original author_portion as org name (before und-splitting)
                full_author = re.sub(r'\s+[A-Z]\.?\s*$', '', author_portion).strip()
                if full_author != primary_author:
                    citation_info = {
                        "authors": [{"family": full_author, "given": ""}],
                        "year": year
                    }
                    result = self.find_match(citation_info)

            if not result and len(words) > 2:
                # Might be a title-like citation e.g. "Advances in Molecular Biology 2014"
                citation_info = {
                    "title": author_portion,
                    "year": year,
                    "authors": []
                }
                result = self.find_match(citation_info)

            if result:
                matches.append(result)

        return matches


# ── Zotero Field Code Generation ──────────────

def zotero_item_to_csl_json(item):
    data = item.get("data", {})
    type_map = {
        "journalArticle": "article-journal", "book": "book",
        "bookSection": "chapter", "conferencePaper": "paper-conference",
        "thesis": "thesis", "report": "report", "webpage": "webpage",
        "magazineArticle": "article-magazine", "newspaperArticle": "article-newspaper",
        "manuscript": "manuscript", "document": "article",
    }

    csl = {"id": item.get("key", ""), "type": type_map.get(data.get("itemType", ""), "article")}

    if data.get("title"):
        csl["title"] = data["title"]

    authors, editors = [], []
    for creator in data.get("creators", []):
        person = {}
        if creator.get("lastName"):
            person["family"] = creator["lastName"]
        if creator.get("firstName"):
            person["given"] = creator["firstName"]
        if creator.get("name"):
            person["literal"] = creator["name"]
        if creator.get("creatorType") == "editor":
            editors.append(person)
        else:
            authors.append(person)

    if authors:
        csl["author"] = authors
    if editors:
        csl["editor"] = editors

    if data.get("date"):
        year_match = re.search(r'(\d{4})', data["date"])
        if year_match:
            csl["issued"] = {"date-parts": [[year_match.group(1)]]}

    for field, csl_field in [
        ("publicationTitle", "container-title"),
        ("bookTitle", "container-title"),
        ("proceedingsTitle", "container-title"),
    ]:
        if data.get(field):
            csl[csl_field] = data[field]
            break

    for zot_field, csl_field in {
        "volume": "volume", "issue": "issue", "pages": "page",
        "publisher": "publisher", "place": "publisher-place",
        "DOI": "DOI", "ISBN": "ISBN", "ISSN": "ISSN",
        "url": "URL", "edition": "edition",
    }.items():
        if data.get(zot_field):
            csl[csl_field] = data[zot_field]

    return csl


def build_zotero_citation_json(zotero_items, user_id):
    citation_id = str(uuid.uuid4()).replace("-", "")[:12]
    citation_items = []
    for item in zotero_items:
        csl_data = zotero_item_to_csl_json(item)
        item_key = item.get("key", "")
        citation_items.append({
            "id": item_key,
            "uris": [f"http://zotero.org/users/{user_id}/items/{item_key}"],
            "uri": [f"http://zotero.org/users/{user_id}/items/{item_key}"],
            "itemData": csl_data,
        })

    citation = {
        "citationID": citation_id,
        "properties": {
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }
    # Do NOT include formattedCitation / plainCitation in properties.
    # When Zotero sees a field code without formattedCitation, it treats
    # the citation as new and regenerates the display text on first refresh
    # without showing the "you have modified this citation" warning.
    return citation


# ── XML Manipulation ──────────────────────────

def create_zotero_field_xml(citation_json_str, display_text):
    W = f"{{{NAMESPACES['w']}}}"
    runs = []

    r_begin = etree.Element(f"{W}r")
    fc_begin = etree.SubElement(r_begin, f"{W}fldChar")
    fc_begin.set(f"{W}fldCharType", "begin")
    runs.append(r_begin)

    instr_text = f" ADDIN ZOTERO_ITEM CSL_CITATION {citation_json_str} "
    chunk_size = 250
    for chunk in [instr_text[i:i+chunk_size] for i in range(0, len(instr_text), chunk_size)]:
        r_instr = etree.Element(f"{W}r")
        it = etree.SubElement(r_instr, f"{W}instrText")
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = chunk
        runs.append(r_instr)

    r_sep = etree.Element(f"{W}r")
    fc_sep = etree.SubElement(r_sep, f"{W}fldChar")
    fc_sep.set(f"{W}fldCharType", "separate")
    runs.append(r_sep)

    r_display = etree.Element(f"{W}r")
    rpr = etree.SubElement(r_display, f"{W}rPr")
    etree.SubElement(rpr, f"{W}noProof")
    t = etree.SubElement(r_display, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = display_text
    runs.append(r_display)

    r_end = etree.Element(f"{W}r")
    fc_end = etree.SubElement(r_end, f"{W}fldChar")
    fc_end.set(f"{W}fldCharType", "end")
    runs.append(r_end)

    return runs


def replace_sdt_with_zotero_field(sdt, runs):
    parent = sdt.getparent()
    if parent is None:
        return False

    W = f"{{{NAMESPACES['w']}}}"
    parent_tag = etree.QName(parent.tag).localname if isinstance(parent.tag, str) else ""

    if parent_tag == "p":
        idx = list(parent).index(sdt)
        parent.remove(sdt)
        for i, run in enumerate(runs):
            parent.insert(idx + i, run)
    else:
        idx = list(parent).index(sdt)
        parent.remove(sdt)
        p = etree.Element(f"{W}p")
        for run in runs:
            p.append(run)
        parent.insert(idx, p)

    return True


def create_zotero_bibl_field_xml(style_uri=DEFAULT_STYLE_URI):
    """Create a ZOTERO_BIBL field code to be placed at the end of the document."""
    W = f"{{{NAMESPACES['w']}}}"
    runs = []

    r_begin = etree.Element(f"{W}r")
    fc_begin = etree.SubElement(r_begin, f"{W}fldChar")
    fc_begin.set(f"{W}fldCharType", "begin")
    runs.append(r_begin)

    bibl_json = json.dumps({
        "uncited": [],
        "omitted": [],
        "custom": []
    }, ensure_ascii=False)
    instr_text = f' ADDIN ZOTERO_BIBL {bibl_json} CSL_BIBLIOGRAPHY '

    r_instr = etree.Element(f"{W}r")
    it = etree.SubElement(r_instr, f"{W}instrText")
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = instr_text
    runs.append(r_instr)

    r_sep = etree.Element(f"{W}r")
    fc_sep = etree.SubElement(r_sep, f"{W}fldChar")
    fc_sep.set(f"{W}fldCharType", "separate")
    runs.append(r_sep)

    r_display = etree.Element(f"{W}r")
    t = etree.SubElement(r_display, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = "{Bibliography will be generated by Zotero}"
    runs.append(r_display)

    r_end = etree.Element(f"{W}r")
    fc_end = etree.SubElement(r_end, f"{W}fldChar")
    fc_end.set(f"{W}fldCharType", "end")
    runs.append(r_end)

    return runs


def remove_citavi_bibliography(root):
    """Remove any existing Citavi-generated bibliography from the document.
    Citavi inserts bibliographies as regular paragraphs (not SDTs),
    often preceded by a heading. Returns the number of elements removed."""
    # Look for Citavi bibliography SDTs (tag starting with "CitaviBibliography")
    W_NS = NAMESPACES['w']
    removed = 0
    for sdt in list(root.iter(f"{{{W_NS}}}sdt")):
        sdt_pr = sdt.find(f"{{{W_NS}}}sdtPr")
        if sdt_pr is None:
            continue
        tag_elem = sdt_pr.find(f"{{{W_NS}}}tag")
        if tag_elem is None:
            continue
        tag_val = tag_elem.get(f"{{{W_NS}}}val", "")
        if "CitaviBibliography" in tag_val or "CitaviFilteredBibliography" in tag_val:
            parent = sdt.getparent()
            if parent is not None:
                parent.remove(sdt)
                removed += 1
    return removed


def add_zotero_bibl_at_end(root, style_uri):
    """Append a ZOTERO_BIBL field code as the last paragraph in the document body."""
    W = f"{{{NAMESPACES['w']}}}"
    body = root.find(f"{W}body")
    if body is None:
        return False

    # Check if a Zotero bibliography already exists
    for instr in root.iter(f"{{{NAMESPACES['w']}}}instrText"):
        if instr.text and "ZOTERO_BIBL" in instr.text:
            return False  # Already exists

    # Create a new paragraph with the ZOTERO_BIBL field
    p = etree.Element(f"{W}p")
    bibl_runs = create_zotero_bibl_field_xml(style_uri)
    for run in bibl_runs:
        p.append(run)

    # Insert before the last sectPr if present, otherwise append to body
    sect_pr = body.find(f"{W}sectPr")
    if sect_pr is not None:
        body.insert(list(body).index(sect_pr), p)
    else:
        body.append(p)

    return True


def verify_document_integrity(root_or_path, log_callback=None):
    """Verify the XML integrity of the document after modification.
    Accepts either an lxml root element or a file path."""
    def log(msg):
        if log_callback:
            log_callback(msg)

    issues = []

    if isinstance(root_or_path, str):
        try:
            parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            tree = etree.parse(root_or_path, parser)
            root = tree.getroot()
        except etree.XMLSyntaxError as e:
            issues.append(f"XML syntax error: {e}")
            return issues
    else:
        root = root_or_path

    W_NS = NAMESPACES['w']

    # Check 1: All fldChar elements are properly paired (begin/separate/end)
    field_stack = 0
    for elem in root.iter(f"{{{W_NS}}}fldChar"):
        fld_type = elem.get(f"{{{W_NS}}}fldCharType", "")
        if fld_type == "begin":
            field_stack += 1
        elif fld_type == "end":
            field_stack -= 1
        if field_stack < 0:
            issues.append("Unmatched field 'end' without 'begin'")
            break

    if field_stack > 0:
        issues.append(f"{field_stack} unclosed field(s) detected (begin without end)")

    # Check 2: Every paragraph is inside the body
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        issues.append("Document body element not found")

    # Check 3: No orphaned SDT content
    for sdt in root.iter(f"{{{W_NS}}}sdt"):
        sdt_content = sdt.find(f"{{{W_NS}}}sdtContent")
        if sdt_content is None:
            issues.append("SDT element without sdtContent found")

    if not issues:
        log("  Document integrity check: PASSED")
    else:
        for issue in issues:
            log(f"  Integrity issue: {issue}")

    return issues


# ── Conversion Engine ─────────────────────────

def process_xml_file(xml_path, matcher, user_id, stats, log_callback=None,
                     style_uri=DEFAULT_STYLE_URI):
    def log(msg):
        if log_callback:
            log_callback(msg)

    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    tree = etree.parse(xml_path, parser)
    root = tree.getroot()
    citavi_sdts = find_citavi_sdts(root)

    if not citavi_sdts:
        return False

    log(f"  Found {len(citavi_sdts)} Citavi citation(s) in {os.path.basename(xml_path)}")
    modified = False

    for sdt, tag_val in citavi_sdts:
        display_text = extract_citavi_display_text(sdt)
        log(f"\n  Processing: {display_text or '[no display text]'}")

        payload = decode_citavi_payload(sdt)
        citation_infos = extract_citation_info_from_citavi(payload) if payload else []

        if payload is None:
            log(f"    Warning: Could not decode Citavi payload (will try display text)")
        elif citation_infos:
            log(f"    Decoded {len(citation_infos)} citation(s) from payload")

        zotero_items = []
        for ci in citation_infos:
            match = matcher.find_match(ci)
            if match:
                zotero_items.append(match)
                log(f"    Matched: {match['data'].get('title', '?')[:60]}")
            else:
                author_str = ci.get("authors", [{}])[0].get("family", "?") if ci.get("authors") else "?"
                log(f"    No match: {author_str}, {ci.get('year', '?')}")
                stats["unmatched"].append(f"{author_str}, {ci.get('year', '?')}")

        if not zotero_items and display_text:
            log(f"    Trying display text fallback...")
            zotero_items = matcher.find_match_by_display_text(display_text)
            if zotero_items:
                log(f"    Matched {len(zotero_items)} item(s) via display text")

        if not zotero_items:
            log(f"    SKIPPED - no Zotero match")
            stats["skipped"] += 1
            continue

        citation_json = build_zotero_citation_json(zotero_items, user_id)
        citation_json_str = json.dumps(citation_json, ensure_ascii=False)
        # Keep the original Citavi display text as visible text in the field.
        # Since we omit formattedCitation from the JSON, Zotero will treat this
        # as a new citation and regenerate display text on first refresh
        # without the "you have modified this citation" warning.
        final_display = display_text if display_text else "(Citation)"
        runs = create_zotero_field_xml(citation_json_str, final_display)

        if replace_sdt_with_zotero_field(sdt, runs):
            stats["converted"] += 1
            modified = True
            log(f"    Replaced with Zotero field code")

    # Only for the main document.xml: handle bibliography
    if os.path.basename(xml_path) == "document.xml":
        # Remove any Citavi bibliography SDTs
        bib_removed = remove_citavi_bibliography(root)
        if bib_removed > 0:
            log(f"\n  Removed {bib_removed} Citavi bibliography element(s)")
            modified = True

        # Add Zotero bibliography field at the end
        if stats.get("converted", 0) > 0 or modified:
            if add_zotero_bibl_at_end(root, style_uri):
                log(f"  Added Zotero bibliography field at end of document")
                modified = True

    if modified:
        # Verify document integrity before writing
        verify_document_integrity(root, log_callback)
        tree.write(xml_path, xml_declaration=True, encoding="UTF-8", standalone=True)

    return modified


def run_conversion(input_path, output_path, library_id, api_key, library_type,
                   style_uri=DEFAULT_STYLE_URI,
                   log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)

    stats = {"converted": 0, "skipped": 0, "unmatched": []}

    log("Connecting to Zotero API...")
    try:
        matcher = ZoteroMatcher(library_id, api_key, library_type)
        count = matcher.get_item_count()
        log(f"Found {count} items in Zotero library.")
    except Exception as e:
        # Sanitize error to avoid leaking API key in logs
        err_msg = str(e)
        if api_key and len(api_key) > 4:
            err_msg = err_msg.replace(api_key, api_key[:4] + "****")
        log(f"ERROR: Could not connect to Zotero: {err_msg}")
        return stats

    log(f"Citation style: {style_uri}")

    with tempfile.TemporaryDirectory() as tmpdir:
        extract_dir = os.path.join(tmpdir, "docx_contents")

        log("\nExtracting .docx...")
        with zipfile.ZipFile(input_path, "r") as z:
            # Validate all member paths to prevent zip-slip path traversal
            for member in z.namelist():
                member_path = os.path.realpath(os.path.join(extract_dir, member))
                if not member_path.startswith(os.path.realpath(extract_dir)):
                    raise ValueError(
                        f"Unsafe path in .docx archive: {member}. "
                        "The file may be malicious."
                    )
            z.extractall(extract_dir)

        for xml_name in ["document.xml", "footnotes.xml", "endnotes.xml"]:
            xml_path = os.path.join(extract_dir, "word", xml_name)
            if os.path.exists(xml_path):
                log(f"\nProcessing word/{xml_name}...")
                process_xml_file(xml_path, matcher, library_id, stats,
                                log_callback, style_uri)

        log("\nRepacking .docx...")
        content_types_path = os.path.join(extract_dir, "[Content_Types].xml")
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            # Write [Content_Types].xml first (OOXML spec requirement)
            if os.path.exists(content_types_path):
                zout.write(content_types_path, "[Content_Types].xml",
                          compress_type=zipfile.ZIP_STORED)
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file == "[Content_Types].xml" and root_dir == extract_dir:
                        continue  # Already written first
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zout.write(file_path, arcname)

    return stats


# ── Verification ──────────────────────────────

def verify_conversion(original_path, converted_path, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("\n" + "=" * 50)
    log("VERIFICATION")
    log("=" * 50)

    results = {
        "original_citavi_count": 0,
        "converted_citavi_remaining": 0,
        "converted_zotero_count": 0,
        "success": False,
    }

    def count_fields_in_docx(docx_path, field_type):
        count = 0
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(docx_path, "r") as z:
                # Validate paths to prevent zip-slip
                for member in z.namelist():
                    member_path = os.path.realpath(os.path.join(tmpdir, member))
                    if not member_path.startswith(os.path.realpath(tmpdir)):
                        raise ValueError(f"Unsafe path in archive: {member}")
                z.extractall(tmpdir)

            for xml_name in ["document.xml", "footnotes.xml", "endnotes.xml"]:
                xml_path = os.path.join(tmpdir, "word", xml_name)
                if not os.path.exists(xml_path):
                    continue

                parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
                tree = etree.parse(xml_path, parser)

                if field_type == "citavi":
                    sdts = find_citavi_sdts(tree)
                    count += len(sdts)
                elif field_type == "zotero":
                    root = tree.getroot()
                    for instr in root.iter(f"{{{NAMESPACES['w']}}}instrText"):
                        if instr.text and "ZOTERO_ITEM" in instr.text:
                            count += 1

        return count

    orig_citavi = count_fields_in_docx(original_path, "citavi")
    results["original_citavi_count"] = orig_citavi
    log(f"  Original: {orig_citavi} Citavi citation(s)")

    conv_citavi = count_fields_in_docx(converted_path, "citavi")
    results["converted_citavi_remaining"] = conv_citavi
    log(f"  Converted: {conv_citavi} Citavi citation(s) remaining")

    conv_zotero = count_fields_in_docx(converted_path, "zotero")
    results["converted_zotero_count"] = conv_zotero
    log(f"  Converted: {conv_zotero} Zotero citation(s)")

    if orig_citavi == 0:
        log("\n  No Citavi citations found in original.")
    elif conv_citavi == 0 and conv_zotero > 0:
        log(f"\n  All {orig_citavi} Citavi citations successfully converted!")
        results["success"] = True
    elif conv_zotero > 0:
        log(f"\n  Partial: {conv_zotero} converted, {conv_citavi} Citavi remaining.")
    else:
        log("\n  Conversion may have failed.")

    return results


# ── AppleScript / Word Automation ─────────────

def open_in_word_and_refresh(original_path, converted_path, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)

    if platform.system() != "Darwin":
        log("Word automation is only available on macOS.")
        return False

    log("\nOpening documents in Microsoft Word...")

    # Sanitize file paths to prevent AppleScript injection.
    # Reject paths containing characters that could break out of
    # the POSIX file context in AppleScript.
    for path_check in [original_path, converted_path]:
        if '"' in path_check or '\\' in path_check or '\'' in path_check or '\n' in path_check or '\r' in path_check:
            log("Warning: File paths contain unsafe characters. Falling back to 'open' command.")
            subprocess.run(['open', converted_path], check=False)
            subprocess.run(['open', original_path], check=False)
            return False

    script = (
        'tell application "Microsoft Word"\n'
        '    activate\n'
        f'    set origDoc to open file name POSIX file "{original_path}"\n'
        f'    set convDoc to open file name POSIX file "{converted_path}"\n'
        '    delay 2\n'
        '    try\n'
        '        tell convDoc\n'
        '            update every field of convDoc\n'
        '        end tell\n'
        '    end try\n'
        '    try\n'
        '        save convDoc\n'
        '    end try\n'
        'end tell\n'
    )

    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            log("Documents opened in Word successfully.")
            log("NOTE: For full refresh, open Zotero and click 'Refresh' in Word toolbar.")
            return True
        else:
            log(f"Word automation warning: {result.stderr.strip()}")
            subprocess.run(['open', converted_path], check=False)
            subprocess.run(['open', original_path], check=False)
            return False
    except (subprocess.TimeoutExpired, FileNotFoundError):
        log("Could not automate Word. Opening files directly...")
        subprocess.run(['open', converted_path], check=False)
        subprocess.run(['open', original_path], check=False)
        return False


# ══════════════════════════════════════════════
# WORKER THREAD (PyQt6 QThread)
# ══════════════════════════════════════════════

class ConversionWorker(QThread):
    """Background thread for running the conversion."""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)

    def __init__(self, input_path, library_id, api_key, library_type,
                 open_word, verify, style_uri=DEFAULT_STYLE_URI):
        super().__init__()
        self.input_path = input_path
        self.library_id = library_id
        self.api_key = api_key
        self.library_type = library_type
        self.open_word = open_word
        self.verify = verify
        self.style_uri = style_uri

    def _log(self, msg):
        self.log_signal.emit(msg)

    def run(self):
        try:
            input_path = self.input_path
            self._log("=" * 50)
            self._log("CITEMIGRATE — CITATION CONVERSION")
            self._log("=" * 50)

            # Step 1: Create copy (use O_EXCL via fd for atomic creation)
            input_dir = os.path.dirname(input_path)
            input_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(input_dir, f"{input_name}_zotero.docx")

            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(input_dir, f"{input_name}_zotero_{counter}.docx")
                counter += 1
                if counter > 1000:
                    raise RuntimeError("Too many output files — clean up old _zotero files.")

            self._log(f"\nOriginal:  {os.path.basename(input_path)}")
            self._log(f"Output:    {os.path.basename(output_path)}")
            self._log("\nStep 1: Creating copy of original document...")
            shutil.copy2(input_path, output_path)
            self._log(f"  Copy created: {os.path.basename(output_path)}")

            # Step 2: Convert
            self._log("\nStep 2: Converting citations...")
            stats = run_conversion(
                output_path, output_path,
                self.library_id, self.api_key, self.library_type,
                style_uri=self.style_uri,
                log_callback=self._log
            )

            self._log(f"\n--- Conversion Results ---")
            self._log(f"  Converted:  {stats['converted']}")
            self._log(f"  Skipped:    {stats['skipped']}")
            self._log(f"  Unmatched:  {len(stats['unmatched'])}")

            if stats["unmatched"]:
                self._log(f"\n  Unmatched references:")
                for ref in stats["unmatched"]:
                    self._log(f"    - {ref}")

            # Step 3: Open in Word
            if self.open_word:
                self._log("\nStep 3: Opening in Word...")
                open_in_word_and_refresh(input_path, output_path, log_callback=self._log)
            else:
                self._log("\nStep 3: Skipped (Word opening disabled)")

            # Step 4: Verify
            verify_results = {}
            if self.verify:
                self._log("\nStep 4: Verifying conversion...")
                verify_results = verify_conversion(
                    input_path, output_path, log_callback=self._log
                )

            result = {
                "stats": stats,
                "verify": verify_results,
                "output_path": output_path,
            }
            self.finished_signal.emit(result)

        except Exception as e:
            # Sanitize error messages to prevent API key leakage in logs
            error_msg = str(e)
            tb = traceback.format_exc()
            if self.api_key and len(self.api_key) > 4:
                masked = self.api_key[:4] + "****"
                error_msg = error_msg.replace(self.api_key, masked)
                tb = tb.replace(self.api_key, masked)
            self._log(f"\nERROR: {error_msg}")
            self._log(tb)
            self.error_signal.emit(error_msg)


class BatchConversionWorker(QThread):
    """Background thread for batch-converting multiple documents."""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)  # current, total

    def __init__(self, input_paths, library_id, api_key, library_type,
                 verify, style_uri=DEFAULT_STYLE_URI):
        super().__init__()
        self.input_paths = input_paths
        self.library_id = library_id
        self.api_key = api_key
        self.library_type = library_type
        self.verify = verify
        self.style_uri = style_uri

    def _log(self, msg):
        self.log_signal.emit(msg)

    def run(self):
        try:
            total = len(self.input_paths)
            self._log("=" * 50)
            self._log(f"CITEMIGRATE — BATCH CONVERSION ({total} files)")
            self._log("=" * 50)

            # Connect to Zotero once (shared matcher for all files)
            self._log("\nConnecting to Zotero API...")
            matcher = ZoteroMatcher(self.library_id, self.api_key, self.library_type)
            count = matcher.get_item_count()
            self._log(f"Found {count} items in Zotero library.")
            self._log(f"Citation style: {self.style_uri}")

            all_stats = {
                "total_files": total,
                "successful": 0,
                "failed": 0,
                "total_converted": 0,
                "total_skipped": 0,
                "total_unmatched": 0,
                "output_paths": [],
            }

            for idx, input_path in enumerate(self.input_paths):
                self.progress_signal.emit(idx + 1, total)
                self._log(f"\n{'─' * 50}")
                self._log(f"File {idx + 1}/{total}: {os.path.basename(input_path)}")
                self._log(f"{'─' * 50}")

                try:
                    input_dir = os.path.dirname(input_path)
                    input_name = os.path.splitext(os.path.basename(input_path))[0]
                    output_path = os.path.join(input_dir, f"{input_name}_zotero.docx")

                    counter = 1
                    while os.path.exists(output_path):
                        output_path = os.path.join(
                            input_dir, f"{input_name}_zotero_{counter}.docx"
                        )
                        counter += 1
                        if counter > 1000:
                            raise RuntimeError(
                                f"Too many output files for {os.path.basename(input_path)}"
                            )

                    shutil.copy2(input_path, output_path)

                    stats = run_conversion(
                        output_path, output_path,
                        self.library_id, self.api_key, self.library_type,
                        style_uri=self.style_uri,
                        log_callback=self._log
                    )

                    all_stats["total_converted"] += stats["converted"]
                    all_stats["total_skipped"] += stats["skipped"]
                    all_stats["total_unmatched"] += len(stats["unmatched"])
                    all_stats["output_paths"].append(output_path)

                    if self.verify:
                        verify_conversion(input_path, output_path, log_callback=self._log)

                    all_stats["successful"] += 1
                    self._log(f"  Result: {stats['converted']} converted, "
                              f"{stats['skipped']} skipped")

                except Exception as file_err:
                    err_msg = str(file_err)
                    if self.api_key and len(self.api_key) > 4:
                        err_msg = err_msg.replace(self.api_key, self.api_key[:4] + "****")
                    self._log(f"  ERROR: {err_msg}")
                    all_stats["failed"] += 1

            self._log(f"\n{'=' * 50}")
            self._log(f"BATCH COMPLETE")
            self._log(f"{'=' * 50}")
            self._log(f"  Files processed: {total}")
            self._log(f"  Successful: {all_stats['successful']}")
            self._log(f"  Failed: {all_stats['failed']}")
            self._log(f"  Total citations converted: {all_stats['total_converted']}")
            self._log(f"  Total skipped: {all_stats['total_skipped']}")

            self.finished_signal.emit(all_stats)

        except Exception as e:
            error_msg = str(e)
            tb = traceback.format_exc()
            if self.api_key and len(self.api_key) > 4:
                masked = self.api_key[:4] + "****"
                error_msg = error_msg.replace(self.api_key, masked)
                tb = tb.replace(self.api_key, masked)
            self._log(f"\nERROR: {error_msg}")
            self._log(tb)
            self.error_signal.emit(error_msg)


# ══════════════════════════════════════════════
# STYLESHEET
# ══════════════════════════════════════════════

STYLESHEET = """
/* ── macOS Native Design Language ─────────────────────
   Follows Apple HIG: system colors, SF font, light appearance,
   subtle borders, rounded corners, native control sizing.
   ──────────────────────────────────────────────────── */
QMainWindow {
    background-color: #f5f5f7;
}
QWidget {
    color: #1d1d1f;
    font-family: -apple-system, "SF Pro Text", "Helvetica Neue", sans-serif;
    font-size: 13px;
}
QGroupBox {
    background-color: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 10px;
    margin-top: 16px;
    padding: 18px 14px 14px 14px;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 6px;
    color: #1d1d1f;
}
QLabel {
    background: transparent;
    border: none;
}
QLineEdit {
    background-color: #ffffff;
    border: 1px solid #c7c7cc;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    color: #1d1d1f;
    selection-background-color: #007aff;
    selection-color: #ffffff;
}
QLineEdit:focus {
    border: 1px solid #007aff;
}
QLineEdit::placeholder {
    color: #8e8e93;
}
QPushButton {
    background-color: #ffffff;
    color: #007aff;
    border: 1px solid #c7c7cc;
    border-radius: 6px;
    padding: 6px 16px;
    font-size: 13px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #f0f0f5;
}
QPushButton:pressed {
    background-color: #e5e5ea;
}
QPushButton:disabled {
    background-color: #f2f2f7;
    color: #c7c7cc;
    border-color: #e5e5ea;
}
QPushButton#convertBtn {
    background-color: #007aff;
    color: #ffffff;
    border: none;
    font-size: 15px;
    font-weight: 600;
    padding: 10px;
    border-radius: 8px;
}
QPushButton#convertBtn:hover {
    background-color: #0066d6;
}
QPushButton#convertBtn:pressed {
    background-color: #004fad;
}
QPushButton#convertBtn:disabled {
    background-color: #c7c7cc;
    color: #ffffff;
}
QComboBox {
    background-color: #ffffff;
    border: 1px solid #c7c7cc;
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 13px;
    color: #1d1d1f;
}
QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 8px;
    color: #1d1d1f;
    selection-background-color: #007aff;
    selection-color: #ffffff;
    padding: 4px;
}
QCheckBox {
    spacing: 8px;
    font-size: 13px;
    padding: 2px 0;
}
QCheckBox::indicator {
    width: 20px;
    height: 20px;
    border: 2px solid #c7c7cc;
    border-radius: 5px;
    background-color: #ffffff;
}
QCheckBox::indicator:hover {
    border-color: #007aff;
}
QTextEdit {
    background-color: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 8px;
    padding: 10px;
    font-family: "SF Mono", "Menlo", "Monaco", monospace;
    font-size: 11px;
    color: #3a3a3c;
    selection-background-color: #007aff;
    selection-color: #ffffff;
}
QProgressBar {
    background-color: #e5e5ea;
    border: none;
    border-radius: 2px;
    height: 4px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #007aff;
    border-radius: 2px;
}
QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #c7c7cc;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #8e8e93;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
"""


# ══════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════

class CiteMigrateApp(QMainWindow):
    # ── Embedded checkmark image (white on transparent, 32x32 PNG) ──
    _CHECKMARK_B64 = (
        "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAf0lEQVR4nO2TQQ7AMAjD"
        "wv7/Z3afWgis9BRfq81upQBCCDGMu3t0/tyQRxFjAV/pLmIkYCUzM7sSUJEfD6jKjwZ0"
        "5FRANqMdjDwNYGa0OmflYQA7o+4LpQGrWzBRlduHAVnECTkAUB8wz9yRA+QMs5935XRA"
        "JPkjb8FOUwghWF6WtFv8WwhlAwAAAABJRU5ErkJggg=="
    )

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CiteMigrate")
        self.setMinimumSize(680, 700)
        self.resize(780, 940)
        self.worker = None
        self._setup_checkmark()
        self._build_ui()
        self._set_app_icon()

    def _setup_checkmark(self):
        """Write checkmark PNG to a temp file for QSS to reference.
        Registers cleanup via atexit to avoid leaving temp files behind."""
        self._checkmark_path = os.path.join(tempfile.gettempdir(), "citemigrate_check.png")
        try:
            with open(self._checkmark_path, "wb") as f:
                f.write(base64.b64decode(self._CHECKMARK_B64))
            # Register cleanup so the temp file is removed on app exit
            atexit.register(
                lambda p=self._checkmark_path: os.unlink(p) if os.path.exists(p) else None
            )
        except Exception:
            self._checkmark_path = ""

    def _set_app_icon(self):
        """Set app icon. Uses icon.png next to the script if available,
        otherwise tries a bundled resource path (for PyInstaller builds)."""
        icon_candidates = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png"),
            os.path.join(getattr(sys, '_MEIPASS', ''), "icon.png"),
        ]
        for icon_path in icon_candidates:
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                return

    def _build_ui(self):
        # ── Scroll area wrapping everything for proper scaling ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("QScrollArea { background-color: #f5f5f7; border: none; }")
        self.setCentralWidget(scroll)

        central = QWidget()
        scroll.setWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # ── Header ────────────────────────────
        header_row = QHBoxLayout()
        header_row.setSpacing(12)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title = QLabel("CiteMigrate")
        title.setFont(QFont("-apple-system", 22, QFont.Weight.Bold))
        title.setStyleSheet("color: #1d1d1f; margin-bottom: 0px;")
        title_col.addWidget(title)

        subtitle = QLabel("Convert Citavi® citations in Word documents to Zotero® field codes")
        subtitle.setStyleSheet("color: #8e8e93; font-size: 12px;")
        title_col.addWidget(subtitle)

        disclaimer = QLabel(
            '<a href="#" style="color: #007aff; font-size: 11px;">'
            'Legal Notice &amp; Disclaimer</a>'
        )
        disclaimer.setCursor(Qt.CursorShape.PointingHandCursor)
        disclaimer.linkActivated.connect(self._show_disclaimer)
        title_col.addWidget(disclaimer)

        header_row.addLayout(title_col, stretch=1)
        layout.addLayout(header_row)

        # ── Tabs: Single File / Batch ─────────
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: none; }
            QTabBar::tab {
                background: #e5e5ea; color: #3a3a3c;
                padding: 6px 20px; border-radius: 6px;
                margin-right: 4px; font-size: 13px;
            }
            QTabBar::tab:selected {
                background: #007aff; color: #ffffff; font-weight: 600;
            }
        """)

        # ── Tab 1: Single File ────────────────
        single_tab = QWidget()
        single_layout = QVBoxLayout(single_tab)
        single_layout.setContentsMargins(0, 12, 0, 0)
        single_layout.setSpacing(10)

        file_group = QGroupBox("Word Document")
        file_gl = QHBoxLayout(file_group)
        file_gl.setContentsMargins(14, 22, 14, 14)
        file_gl.setSpacing(8)
        self.file_entry = QLineEdit()
        self.file_entry.setPlaceholderText("Select a .docx file...")
        self.file_entry.setMinimumHeight(32)
        file_gl.addWidget(self.file_entry)
        browse_btn = QPushButton("Browse")
        browse_btn.setFixedWidth(90)
        browse_btn.setMinimumHeight(32)
        browse_btn.clicked.connect(self._pick_file)
        file_gl.addWidget(browse_btn)
        single_layout.addWidget(file_group)

        self.tabs.addTab(single_tab, "  Single File  ")

        # ── Tab 2: Batch ──────────────────────
        batch_tab = QWidget()
        batch_layout = QVBoxLayout(batch_tab)
        batch_layout.setContentsMargins(0, 12, 0, 0)
        batch_layout.setSpacing(10)

        batch_group = QGroupBox("Word Documents (Batch)")
        batch_gl = QVBoxLayout(batch_group)
        batch_gl.setContentsMargins(14, 22, 14, 14)
        batch_gl.setSpacing(8)

        self.batch_list = QListWidget()
        self.batch_list.setMinimumHeight(100)
        self.batch_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.batch_list.setStyleSheet(
            "QListWidget { background: #ffffff; border: 1px solid #d2d2d7; "
            "border-radius: 6px; padding: 4px; font-size: 12px; }"
            "QListWidget::item { padding: 4px 6px; border-radius: 4px; }"
            "QListWidget::item:selected { background: #007aff; color: white; }"
        )
        batch_gl.addWidget(self.batch_list)

        batch_btn_row = QHBoxLayout()
        batch_btn_row.setSpacing(8)
        add_files_btn = QPushButton("Add Files")
        add_files_btn.setMinimumHeight(30)
        add_files_btn.clicked.connect(self._add_batch_files)
        batch_btn_row.addWidget(add_files_btn)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.setMinimumHeight(30)
        remove_btn.clicked.connect(self._remove_batch_files)
        batch_btn_row.addWidget(remove_btn)
        clear_btn = QPushButton("Clear All")
        clear_btn.setMinimumHeight(30)
        clear_btn.clicked.connect(lambda: self.batch_list.clear())
        batch_btn_row.addWidget(clear_btn)
        batch_gl.addLayout(batch_btn_row)

        batch_layout.addWidget(batch_group)
        self.tabs.addTab(batch_tab, "  Batch Mode  ")

        layout.addWidget(self.tabs)

        # ── Zotero API Settings ───────────────
        zotero_group = QGroupBox("Zotero API Settings")
        zot_layout = QVBoxLayout(zotero_group)
        zot_layout.setSpacing(8)
        zot_layout.setContentsMargins(14, 26, 14, 14)

        # Library ID + Library Type on one row
        id_row = QHBoxLayout()
        id_row.setSpacing(12)

        id_col = QVBoxLayout()
        id_col.setSpacing(4)
        id_label = QLabel("Library ID")
        id_label.setStyleSheet("color: #6e6e73; font-size: 12px; font-weight: 500;")
        id_col.addWidget(id_label)
        self.library_id_entry = QLineEdit()
        self.library_id_entry.setPlaceholderText("e.g. 12345678")
        self.library_id_entry.setMinimumHeight(34)
        id_col.addWidget(self.library_id_entry)
        id_row.addLayout(id_col, stretch=3)

        type_col = QVBoxLayout()
        type_col.setSpacing(4)
        type_label = QLabel("Library Type")
        type_label.setStyleSheet("color: #6e6e73; font-size: 12px; font-weight: 500;")
        type_col.addWidget(type_label)
        self.library_type_combo = QComboBox()
        self.library_type_combo.addItems(["user", "group"])
        self.library_type_combo.setMinimumHeight(34)
        type_col.addWidget(self.library_type_combo)
        id_row.addLayout(type_col, stretch=1)

        zot_layout.addLayout(id_row)

        # API Key – full width with more space
        key_col = QVBoxLayout()
        key_col.setSpacing(4)
        key_label = QLabel("API Key")
        key_label.setStyleSheet("color: #6e6e73; font-size: 12px; font-weight: 500; margin-top: 4px;")
        key_col.addWidget(key_label)
        self.api_key_entry = QLineEdit()
        self.api_key_entry.setPlaceholderText("Your Zotero API key")
        self.api_key_entry.setMinimumHeight(34)
        self.api_key_entry.setEchoMode(QLineEdit.EchoMode.Password)
        key_col.addWidget(self.api_key_entry)
        zot_layout.addLayout(key_col)

        hint = QLabel(
            '<span style="color: #8e8e93; font-size: 11px;">'
            'Get your API key at '
            '<a href="https://www.zotero.org/settings/keys" style="color: #007aff;">zotero.org/settings/keys</a>'
            '  ·  Library ID is your userID shown on the same page</span>'
        )
        hint.setOpenExternalLinks(True)
        hint.setWordWrap(True)
        zot_layout.addWidget(hint)

        layout.addWidget(zotero_group)

        # ── Citation Style ────────────────────
        style_group = QGroupBox("Citation Style")
        style_layout = QVBoxLayout(style_group)
        style_layout.setSpacing(8)
        style_layout.setContentsMargins(14, 22, 14, 14)

        self.style_combo = QComboBox()
        self.style_combo.setMinimumHeight(32)
        self._citation_styles = {
            "Harvard - Cite Them Right": DEFAULT_STYLE_URI,
            "APA 7th Edition": "http://www.zotero.org/styles/apa",
            "APA 6th Edition": "http://www.zotero.org/styles/apa-6th-edition",
            "Chicago (Author-Date)": "http://www.zotero.org/styles/chicago-author-date",
            "Chicago (Note)": "http://www.zotero.org/styles/chicago-note-bibliography",
            "Vancouver": "http://www.zotero.org/styles/vancouver",
            "IEEE": "http://www.zotero.org/styles/ieee",
            "Nature": "http://www.zotero.org/styles/nature",
            "BMJ": "http://www.zotero.org/styles/bmj",
            "AMA 11th Edition": "http://www.zotero.org/styles/american-medical-association",
            "MLA 9th Edition": "http://www.zotero.org/styles/modern-language-association",
            "Springer - Basic (Author-Date)": "http://www.zotero.org/styles/springer-basic-author-date",
            "Elsevier - Harvard": "http://www.zotero.org/styles/elsevier-harvard",
            "DIN 1505-2": "http://www.zotero.org/styles/din-1505-2",
        }
        for name in self._citation_styles:
            self.style_combo.addItem(name)
        self.style_combo.setCurrentIndex(0)
        style_layout.addWidget(self.style_combo)

        custom_style_row = QHBoxLayout()
        self.custom_style_cb = QCheckBox("Custom style URI:")
        self.custom_style_entry = QLineEdit()
        self.custom_style_entry.setPlaceholderText("http://www.zotero.org/styles/your-style")
        self.custom_style_entry.setMinimumHeight(30)
        self.custom_style_entry.setEnabled(False)
        self.custom_style_cb.toggled.connect(
            lambda checked: (
                self.custom_style_entry.setEnabled(checked),
                self.style_combo.setEnabled(not checked),
            )
        )
        custom_style_row.addWidget(self.custom_style_cb)
        custom_style_row.addWidget(self.custom_style_entry, stretch=1)
        style_layout.addLayout(custom_style_row)

        layout.addWidget(style_group)

        # ── Options ───────────────────────────
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout(options_group)
        options_layout.setSpacing(10)
        options_layout.setContentsMargins(14, 22, 14, 14)

        self.open_word_cb = QCheckBox("  Open documents in Word after conversion (single file only)")
        self.open_word_cb.setChecked(True)
        options_layout.addWidget(self.open_word_cb)

        self.verify_cb = QCheckBox("  Verify conversion results")
        self.verify_cb.setChecked(True)
        options_layout.addWidget(self.verify_cb)

        layout.addWidget(options_group)

        # ── Convert Button ────────────────────
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setObjectName("convertBtn")
        self.convert_btn.setMinimumHeight(44)
        self.convert_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_btn.clicked.connect(self._start_conversion)
        layout.addWidget(self.convert_btn)

        # ── Progress Bar ──────────────────────
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setRange(0, 0)  # indeterminate
        self.progress.setVisible(False)
        self.progress.setTextVisible(False)
        layout.addWidget(self.progress)

        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("color: #8e8e93; font-size: 11px;")
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)

        # ── Log Area ──────────────────────────
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(160)
        layout.addWidget(self.log_text, stretch=1)

    # ── Helpers ───────────────────────────────

    def _show_disclaimer(self, _link=None):
        QMessageBox.information(
            self, "Legal Notice & Disclaimer",
            "CiteMigrate — Legal Notice\n"
            "═══════════════════════════════════════\n\n"
            "DISCLAIMER\n"
            "This software is provided \"as is\" without warranty of any kind. "
            "Always test with a copy of your document and verify converted "
            "citations before using in important work.\n\n"
            "TRADEMARKS\n"
            "\"Citavi\" is a trademark of Swiss Academic Software GmbH.\n"
            "\"Zotero\" is a registered trademark of the Corporation for "
            "Digital Scholarship.\n"
            "\"Microsoft Word\" is a trademark of Microsoft Corporation.\n\n"
            "This tool is not affiliated with, endorsed by, or produced by "
            "any of these organizations. Trademark names are used solely "
            "for descriptive purposes to identify compatibility.\n\n"
            "TECHNICAL APPROACH\n"
            "This tool reads the publicly documented Office Open XML (OOXML / "
            "ECMA-376) format of .docx files. No proprietary software from "
            "Citavi, Zotero, or Microsoft was decompiled, disassembled, or "
            "reverse-engineered. The tool only parses the openly accessible "
            "XML structure within .docx documents and communicates with Zotero "
            "through its official public API.\n\n"
            "AI-GENERATED SOFTWARE\n"
            "This software was developed with the assistance of Claude "
            "(Anthropic) through iterative human-AI collaboration.\n\n"
            "LICENSE\n"
            "Licensed under the GNU General Public License v3.0 (GPLv3).\n"
            "See the LICENSE file for full details."
        )

    def _pick_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Word Document", "",
            "Word Documents (*.docx);;All Files (*)"
        )
        if filename:
            self.file_entry.setText(filename)

    def _add_batch_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Word Documents", "",
            "Word Documents (*.docx);;All Files (*)"
        )
        for f in files:
            # Avoid duplicates
            existing = [self.batch_list.item(i).data(Qt.ItemDataRole.UserRole)
                        for i in range(self.batch_list.count())]
            if f not in existing:
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.ItemDataRole.UserRole, f)
                item.setToolTip(f)
                self.batch_list.addItem(item)

    def _remove_batch_files(self):
        for item in self.batch_list.selectedItems():
            self.batch_list.takeItem(self.batch_list.row(item))

    def _log(self, message):
        self.log_text.append(message)
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _get_style_uri(self):
        if self.custom_style_cb.isChecked() and self.custom_style_entry.text().strip():
            return self.custom_style_entry.text().strip()
        style_name = self.style_combo.currentText()
        return self._citation_styles.get(
            style_name, DEFAULT_STYLE_URI
        )

    def _validate_api_inputs(self):
        if not self.library_id_entry.text().strip():
            QMessageBox.warning(self, "Missing Library ID", "Please enter your Zotero Library ID.")
            return False
        if not self.api_key_entry.text().strip():
            QMessageBox.warning(self, "Missing API Key", "Please enter your Zotero API Key.")
            return False
        return True

    # ── Conversion ────────────────────────────

    def _start_conversion(self):
        if not self._validate_api_inputs():
            return

        is_batch = self.tabs.currentIndex() == 1

        if is_batch:
            # Batch mode
            if self.batch_list.count() == 0:
                QMessageBox.warning(self, "No Files", "Please add at least one .docx file.")
                return
            input_paths = [
                self.batch_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.batch_list.count())
            ]
            # Validate all files exist
            for p in input_paths:
                if not os.path.exists(p):
                    QMessageBox.warning(self, "File Not Found", f"File not found:\n{p}")
                    return

            self.log_text.clear()
            self.convert_btn.setEnabled(False)
            self.convert_btn.setText(f"Converting {len(input_paths)} files...")
            self.progress.setVisible(True)
            self.progress.setRange(0, len(input_paths))
            self.progress.setValue(0)
            self.progress_label.setVisible(True)
            self.progress_label.setText(f"0 / {len(input_paths)}")

            self.worker = BatchConversionWorker(
                input_paths=input_paths,
                library_id=self.library_id_entry.text().strip(),
                api_key=self.api_key_entry.text().strip(),
                library_type=self.library_type_combo.currentText(),
                verify=self.verify_cb.isChecked(),
                style_uri=self._get_style_uri(),
            )
            self.worker.log_signal.connect(self._log)
            self.worker.progress_signal.connect(self._on_batch_progress)
            self.worker.finished_signal.connect(self._on_batch_finished)
            self.worker.error_signal.connect(self._on_error)
            self.worker.start()
        else:
            # Single file mode
            file_path = self.file_entry.text().strip()
            if not file_path:
                QMessageBox.warning(self, "Missing File", "Please select a Word document.")
                return
            if not os.path.exists(file_path):
                QMessageBox.warning(self, "File Not Found", f"File not found:\n{file_path}")
                return
            if not file_path.lower().endswith(".docx"):
                QMessageBox.warning(self, "Invalid File", "Please select a .docx file.")
                return

            self.log_text.clear()
            self.convert_btn.setEnabled(False)
            self.convert_btn.setText("Converting...")
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)  # indeterminate
            self.progress_label.setVisible(False)

            self.worker = ConversionWorker(
                input_path=file_path,
                library_id=self.library_id_entry.text().strip(),
                api_key=self.api_key_entry.text().strip(),
                library_type=self.library_type_combo.currentText(),
                open_word=self.open_word_cb.isChecked(),
                verify=self.verify_cb.isChecked(),
                style_uri=self._get_style_uri(),
            )
            self.worker.log_signal.connect(self._log)
            self.worker.finished_signal.connect(self._on_finished)
            self.worker.error_signal.connect(self._on_error)
            self.worker.start()

    def _cleanup_worker(self):
        if self.worker is not None:
            self.worker.quit()
            if not self.worker.wait(5000):
                # Thread didn't finish in 5s — force terminate
                self.worker.terminate()
                self.worker.wait(2000)
            self.worker = None

    def _on_finished(self, result):
        self._cleanup_worker()
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert")
        self.progress.setVisible(False)

        stats = result.get("stats", {})
        verify = result.get("verify", {})
        output = result.get("output_path", "")

        if verify.get("success"):
            QMessageBox.information(
                self, "Success",
                f"Conversion complete!\n\n"
                f"{verify['original_citavi_count']} Citavi citations "
                f"converted to {verify['converted_zotero_count']} Zotero citations.\n\n"
                f"Output: {os.path.basename(output)}"
            )
        elif stats.get("converted", 0) > 0:
            QMessageBox.warning(
                self, "Conversion Complete",
                f"Conversion finished with warnings.\n\n"
                f"Converted: {stats['converted']}\n"
                f"Skipped: {stats['skipped']}\n"
                f"Unmatched: {len(stats.get('unmatched', []))}\n\n"
                f"Check the log for details."
            )
        else:
            QMessageBox.warning(
                self, "No Conversions",
                "No citations were converted. Check the log for details."
            )

    def _on_batch_progress(self, current, total):
        self.progress.setValue(current)
        self.progress_label.setText(f"{current} / {total}")

    def _on_batch_finished(self, all_stats):
        self._cleanup_worker()
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert")
        self.progress.setVisible(False)
        self.progress_label.setVisible(False)

        QMessageBox.information(
            self, "Batch Complete",
            f"Batch conversion complete!\n\n"
            f"Files processed: {all_stats['total_files']}\n"
            f"Successful: {all_stats['successful']}\n"
            f"Failed: {all_stats['failed']}\n\n"
            f"Total citations converted: {all_stats['total_converted']}\n"
            f"Total skipped: {all_stats['total_skipped']}"
        )

    def _on_error(self, error_msg):
        self._cleanup_worker()
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert")
        self.progress.setVisible(False)
        self.progress_label.setVisible(False)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{error_msg}")

    def closeEvent(self, event):
        """Ensure worker thread is stopped before closing the window."""
        self._cleanup_worker()
        super().closeEvent(event)


# ══════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)

    window = CiteMigrateApp()

    # Inject checkmark image path into checkbox stylesheet (needs runtime path)
    if window._checkmark_path and os.path.exists(window._checkmark_path):
        # Use forward slashes for QSS url() even on Windows
        check_path = window._checkmark_path.replace("\\", "/")
        checkbox_checked_qss = (
            f'QCheckBox::indicator:checked {{\n'
            f'    background-color: #007aff;\n'
            f'    border-color: #007aff;\n'
            f'    image: url("{check_path}");\n'
            f'}}\n'
        )
        app.setStyleSheet(app.styleSheet() + checkbox_checked_qss)
    else:
        # Fallback: solid blue without checkmark image
        app.setStyleSheet(app.styleSheet() +
            'QCheckBox::indicator:checked { background-color: #007aff; border-color: #007aff; }\n')

    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
