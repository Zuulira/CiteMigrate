# CiteMigrate

A desktop application that converts Citavi citation field codes in Microsoft Word documents (.docx) to Zotero-compatible field codes with proper linkage to your Zotero library.

## What it does

If you've migrated your reference library from Citavi to Zotero but still have Word documents with Citavi citation field codes, this tool converts them so Zotero can manage the citations. The original document is never modified — a copy is created with the converted citations.

### Features

- Converts Citavi SDT-based citations to Zotero `ADDIN ZOTERO_ITEM CSL_CITATION` field codes
- Matches citations via DOI, ISBN, title+author+year, and display text fallback
- Supports multilingual citation formats (German, French, Spanish, Italian, Portuguese, Dutch)
- Removes Citavi bibliography and inserts a Zotero bibliography field at the end
- Selectable citation style (Harvard, APA, Chicago, Vancouver, IEEE, and more)
- Single file and batch conversion modes
- Document integrity verification after conversion
- Optional Microsoft Word automation (macOS only) to open and refresh both documents
- Cross-platform GUI (macOS, Windows, Linux) with PyQt6

## Requirements

- macOS, Windows, or Linux
- Python 3.9 or higher
- A Zotero account with your Citavi library already imported
- A Zotero API key ([get one here](https://www.zotero.org/settings/keys))
- Your Zotero Library ID (found at [zotero.org/settings/keys](https://www.zotero.org/settings/keys) under "Your userID for use in API calls")

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/Zuulira/citemigrate.git
cd citemigrate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the application
python citemigrate.py
```

That's it — no compilation or build step needed. The app opens as a GUI window.

### Optional: Build a standalone macOS .app

If you want a self-contained .app bundle that runs without a Python installation:

```bash
pip install pyinstaller
bash build_app.sh
# The app will be in dist/CiteMigrate.app
```

Note: This only works on macOS. The resulting .app is not notarised, so users will need to right-click > Open on first launch to bypass Gatekeeper.

## Usage

1. **Select your Word document** (.docx file with Citavi citations)
2. **Enter your Zotero API credentials** (Library ID, API Key, Library Type)
3. **Choose a citation style** that matches your Zotero configuration
4. **Click Convert**

The tool will:
- Create a copy of your document (e.g., `YourDocument_zotero.docx`)
- Connect to the Zotero API and fetch your library items
- Match each Citavi citation to a Zotero item
- Replace Citavi field codes with Zotero field codes
- Add a Zotero bibliography field at the end of the document
- Optionally open both documents in Word for comparison

**After conversion**, open the converted document in Word with Zotero running and click "Refresh" in the Zotero toolbar to finalize the citations and generate the bibliography.

## How matching works

Citations are matched in the following priority order:

1. **DOI** — exact match
2. **ISBN** — exact match (normalized)
3. **Title + Author + Year** — fuzzy matching with scoring
4. **Display text parsing** — extracts author names and year from the visible citation text as a fallback

The display text parser handles German-specific patterns like "und" as author separator, uppercase names (e.g., "SMITH und JONES 2010"), institutional authors, and citation prefixes like "vgl.".

## Known limitations

- Requires that your Citavi library has already been imported into Zotero
- Citations without sufficient metadata (no DOI, no title, ambiguous display text) may not match
- The Zotero API rate limit may slow down processing for very large libraries
- Word automation (auto-open/refresh) works only on macOS with Microsoft Word installed

## Disclaimer

**This software is provided "as is" without warranty of any kind.** Always test with a copy of your document and verify converted citations before using in important work. Data loss or corruption during conversion is possible.

### Trademark Notice

This tool is **not affiliated with, endorsed by, or produced by** any of the following organizations:

- **Citavi** is a trademark of Swiss Academic Software GmbH.
- **Zotero** is a registered trademark of the Corporation for Digital Scholarship.
- **Microsoft** and **Microsoft Word** are trademarks of Microsoft Corporation.

Trademark names are used in this project solely for descriptive purposes to identify compatibility. No sponsorship or endorsement is implied.

### Technical Approach — No Decompilation

This tool reads .docx files, which use the publicly documented **Office Open XML (OOXML / ECMA-376)** format — an open, standardized file format. The tool parses the openly accessible XML structure within .docx documents to identify Citavi-specific structured document tags (SDTs) and replaces them with Zotero-compatible field codes.

**No proprietary software was decompiled, disassembled, or reverse-engineered.** The tool does not access, modify, or interact with the internal code or binaries of Citavi, Zotero, or Microsoft Word. It communicates with Zotero exclusively through its official, publicly documented [Web API](https://www.zotero.org/support/dev/web_api/v3/start).

### Data & Privacy

- Your Word documents are processed locally on your machine
- When using the Zotero API, your API credentials and citation metadata are transmitted to Zotero's servers per their [Terms of Service](https://www.zotero.org/support/terms/terms_of_service)
- No data is collected, stored, or transmitted to any third party by this tool
- Your Zotero API key is never logged or written to disk by this tool

### AI-Generated Software

This software was developed with the assistance of **Claude** (Anthropic). The code was generated through an iterative human-AI collaboration process. The human author directed the design, tested the output, and verified the results. The AI assisted with code generation, debugging, and documentation.

## License

This project is licensed under the **GNU General Public License v3.0** — see the [LICENSE](LICENSE) file for details.

This license is required because the project depends on PyQt6, which is licensed under GPLv3 for open-source use.

### Third-party components

| Component | License | Purpose |
|-----------|---------|---------|
| [PyQt6](https://riverbankcomputing.com/software/pyqt/) | GPLv3 | GUI framework |
| [lxml](https://lxml.de/) | BSD | XML parsing and manipulation |
| [pyzotero](https://github.com/urschrei/pyzotero) | MIT | Zotero API client |

## Contributing

Contributions are welcome. Please open an issue first to discuss proposed changes.

## Author

Ingo — [GitHub](https://github.com/Zuulira)

Built with assistance from Claude (Anthropic).
