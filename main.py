from flask import Flask, request, render_template
from docx import Document
from pypdf import PdfReader
from spellchecker import SpellChecker
from html import escape
import io
import re

app = Flask(__name__)

# Uploads are read into memory, so cap what one request may carry. Flask
# rejects anything larger with a 413 before the body reaches the route.
MAX_UPLOAD_MB = 10
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024

# Building the checker loads a word-frequency dictionary, so do it once at
# import instead of per request. Lookups are read-only.
spell = SpellChecker()

REQUIRED_HEADINGS = [
    "Purpose",
    "QA Check Sections",
    "Modifications",
    "Initial Inspection",
    "Procedures",
    "Pre Compliance Marks",
    "Compliance Marks",
    "Non-Powered Tests",
    "Initial Assembly",
    "Corner Bonding",
    "Initial Programming",
    "SDK Programming",
    "Top Level Assembly",
    "FPGA JTAG Configuration",
    "Quick Test",
    "Burn-in",
    "Shipping"
]

# Lower-cased required heading -> canonical spelling, so a heading typed
# "PURPOSE" still counts as the required "Purpose".
REQUIRED_LOOKUP = {heading.lower(): heading for heading in REQUIRED_HEADINGS}

# Sections skipped by the spell checker. Test and programming steps are full of
# part numbers, fixture IDs and tool names that are not dictionary words, so
# spellchecking them only produces noise. Edit this list to change the scope.
SPELLCHECK_SKIP_SECTIONS = [
    "Non-Powered Tests",
    "Initial Programming",
    "SDK Programming",
    "FPGA JTAG Configuration",
    "Quick Test",
    "Burn-in",
]

SKIP_KEYS = {name.lower() for name in SPELLCHECK_SKIP_SECTIONS}

# Words that should never ship in a procedure.
FORBIDDEN_WORDS = ["error!", "unauthorized"]

# A heading line is short, unpunctuated and not a sentence. These bounds keep
# body paragraphs and table rows out of the heading list.
MAX_HEADING_WORDS = 9
MAX_HEADING_CHARS = 80

# Leading section numbers: "3", "3.2", "3.2.1", "A.1", optionally with ) or -.
SECTION_NUMBER = re.compile(r"^([0-9]+|[A-Z])(\.[0-9]+)*[.)\-]?\s+")


def strip_section_number(text):
    """"3.2 Purpose " -> "Purpose"."""
    return SECTION_NUMBER.sub("", text.strip()).strip()


def heading_key(heading):
    """Comparison key for a heading: no section number, no case, no padding."""
    return strip_section_number(heading).lower()


def build_heading_lookup(found_headings):
    """Map heading key -> display name for every heading we recognise.

    Covers the headings the document carries plus the required ones, so a
    required heading that was not styled as a heading is still treated as a
    section boundary.
    """
    lookup = dict(REQUIRED_LOOKUP)

    for heading in found_headings:
        name = strip_section_number(heading)
        key = name.lower()

        if key and key not in lookup:
            lookup[key] = name

    return lookup


def extract_docx_data(file_stream):
    doc = Document(io.BytesIO(file_stream.read()))

    full_text = []
    headings = []

    for para in doc.paragraphs:
        text = para.text.strip()

        if not text:
            continue

        full_text.append(text)

        # Detect headings based on style. "Title" and "Subtitle" carry section
        # names in some templates, so they count too.
        style = para.style.name or ""

        if style.startswith('Heading') or style in ('Title', 'Subtitle'):
            headings.append(text)

    return "\n".join(full_text), headings


def looks_like_heading(line):
    """Heuristic heading test for formats with no style information."""
    stripped = strip_section_number(line)

    if not stripped or len(stripped) > MAX_HEADING_CHARS:
        return False

    # Sentences and list/table content are not headings.
    if stripped[-1] in ".,;:!?" or "\t" in line or "|" in stripped:
        return False

    words = stripped.split()

    if not words or len(words) > MAX_HEADING_WORDS:
        return False

    # Needs at least one letter, and no lowercase-only opening word.
    if not any(char.isalpha() for char in stripped):
        return False

    letters = [char for char in stripped if char.isalpha()]

    # ALL CAPS headings.
    if letters and all(char.isupper() for char in letters):
        return True

    # Title Case headings: every word that carries letters starts uppercase,
    # ignoring short connectors like "of", "and", "in".
    connectors = {"a", "an", "and", "at", "by", "for", "in", "of",
                  "on", "or", "the", "to", "with"}

    significant = [w for w in words if any(char.isalpha() for char in w)]

    if not significant or not significant[0][0].isupper():
        return False

    for word in significant[1:]:
        if word.lower() in connectors:
            continue

        first = next(char for char in word if char.isalpha())

        if not first.isupper():
            return False

    return True


def find_headings_in_text(text):
    """Find headings in plain text (PDF and TXT), in document order.

    Required headings are always picked up; anything else that reads like a
    heading is reported too, so the results page can show the full list.
    """
    found = []
    seen = set()

    for line in text.split("\n"):
        if not line.strip():
            continue

        # "1. Purpose" and "Purpose" are the same section.
        bare = strip_section_number(line)
        key = bare.lower()

        if key in REQUIRED_LOOKUP:
            name = REQUIRED_LOOKUP[key]
        elif looks_like_heading(line):
            name = bare
        else:
            continue

        if name.lower() in seen:
            continue

        seen.add(name.lower())
        found.append(name)

    # A required heading may sit inline (run together with body text by the PDF
    # text extractor), so fall back to a substring scan for anything missed.
    lowered = text.lower()

    for heading in REQUIRED_HEADINGS:
        if heading.lower() not in seen and heading.lower() in lowered:
            seen.add(heading.lower())
            found.append(heading)

    return found


def extract_document(uploaded_file):
    """Return (text, headings) for an upload, or (None, None) if unsupported."""
    filename = (uploaded_file.filename or "").lower()

    if filename.endswith('.docx'):
        return extract_docx_data(uploaded_file.stream)

    if filename.endswith('.pdf'):
        reader = PdfReader(uploaded_file.stream)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)

    elif filename.endswith('.txt'):
        text = uploaded_file.stream.read().decode("utf-8", errors="ignore")

    else:
        return None, None

    return text, find_headings_in_text(text)


def tokenize_for_spelling(line):
    """Split a line into checkable words, dropping obvious non-prose tokens."""
    words = []

    for raw in line.split():
        # Part numbers, revisions and measurements (P/N 4832-A, 24VDC, Rev C2).
        if any(char.isdigit() for char in raw):
            continue

        word = raw.strip(".,;:!?()[]{}<>\"'`*/\\|_=+~-").lower()

        # Keep ordinary words only: letters, with internal hyphen or apostrophe.
        if len(word) < 2 or not re.fullmatch(r"[a-z][a-z'-]*", word):
            continue

        words.append(word)

    return words


def check_spelling(text, lookup):
    """Spellcheck the document body, skipping SPELLCHECK_SKIP_SECTIONS."""
    words = []
    skipped = []
    current_key = None

    for line in text.split("\n"):
        key = heading_key(line)

        if key in lookup:
            current_key = key

            if key in SKIP_KEYS and lookup[key] not in skipped:
                skipped.append(lookup[key])

            # Heading text itself is fixed boilerplate -- nothing to check.
            continue

        if current_key in SKIP_KEYS:
            continue

        words.extend(tokenize_for_spelling(line))

    return spell.unknown(words), skipped


def find_forbidden_words(text):
    lowered = text.lower()

    return [word for word in FORBIDDEN_WORDS if word.lower() in lowered]


def check_headings(found_headings):
    """Describe every heading found in the document, in document order.

    Duplicates are collapsed but counted, so a document that repeats
    "Procedures" three times reports it once with a count.
    """
    all_headings = []
    by_key = {}

    for heading in found_headings:
        text = heading.strip()
        key = heading_key(text)

        if not key:
            continue

        if key in by_key:
            by_key[key]["count"] += 1
            continue

        entry = {
            "text": REQUIRED_LOOKUP.get(key, text),
            "required": key in REQUIRED_LOOKUP,
            "count": 1
        }

        by_key[key] = entry
        all_headings.append(entry)

    matched = [h["text"] for h in all_headings if h["required"]]

    return {
        "all_headings": all_headings,
        "found_headings": matched,
        "extra_headings": [h["text"] for h in all_headings if not h["required"]],
        "missing_headings": [h for h in REQUIRED_HEADINGS if h not in matched],
        "required_headings": REQUIRED_HEADINGS
    }


def format_text_with_headings(text, lookup):
    html = ""

    for line in text.split("\n"):
        # Escape document text -- it is rendered into the results page as HTML.
        safe_line = escape(line)

        if heading_key(line) in lookup:
            html += f"<h2>{safe_line}</h2>"
        else:
            html += f"<p>{safe_line}</p>"

    return html


def error_page(message, status):
    return render_template("error.html", message=message), status


@app.errorhandler(413)
def upload_too_large(error):
    return error_page(
        f"That file is larger than the {MAX_UPLOAD_MB} MB limit.", 413)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method != 'POST':
        return render_template("index.html",
                               required_count=len(REQUIRED_HEADINGS))

    uploaded_file = request.files.get('document')

    if not uploaded_file or not uploaded_file.filename:
        return error_page("No file was uploaded.", 400)

    try:
        text, headings = extract_document(uploaded_file)
    except Exception:
        # python-docx and pypdf raise assorted exception types on a file that
        # is corrupt, encrypted or not really the format its name claims.
        app.logger.exception("Could not read %s", uploaded_file.filename)

        return error_page(
            "That file could not be read. It may be corrupt, password "
            "protected, or not really a Word, PDF or text file.", 400)

    if text is None:
        return error_page(
            "Unsupported file type. Upload a .docx, .pdf or .txt file.", 400)

    lookup = build_heading_lookup(headings)
    spelling_errors, skipped_sections = check_spelling(text, lookup)

    return render_template(
        "results.html",
        filename=uploaded_file.filename,
        spelling_errors=spelling_errors,
        skipped_sections=skipped_sections,
        forbidden_words=find_forbidden_words(text),
        formatted_text=format_text_with_headings(text, lookup),
        **check_headings(headings)
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
