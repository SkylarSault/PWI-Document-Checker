from flask import Flask, request, render_template
from docx import Document
from PyPDF2 import PdfReader
from spellchecker import SpellChecker
from html import escape
import io
import re

app = Flask(__name__)

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

# A heading line is short, unpunctuated and not a sentence. These bounds keep
# body paragraphs and table rows out of the heading list.
MAX_HEADING_WORDS = 9
MAX_HEADING_CHARS = 80

# Leading section numbers: "3", "3.2", "3.2.1", "A.1", optionally with ) or -.
SECTION_NUMBER = re.compile(r"^([0-9]+|[A-Z])(\.[0-9]+)*[.)\-]?\s+")


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

def heading_key(heading):
    """Comparison key for a heading: no section number, no case, no padding."""
    return SECTION_NUMBER.sub("", heading.strip()).strip().lower()


def match_heading_line(line, found_headings):
    """Return the canonical heading name if this line is a section heading."""
    key = heading_key(line)

    if not key:
        return None

    for heading in found_headings:
        if heading_key(heading) == key:
            return SECTION_NUMBER.sub("", heading.strip()).strip()

    for heading in REQUIRED_HEADINGS:
        if key == heading.lower():
            return heading

    return None

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

def check_spelling(text, found_headings):
    """Spellcheck the document body, skipping SPELLCHECK_SKIP_SECTIONS."""
    spell = SpellChecker()

    skip_lookup = {name.lower() for name in SPELLCHECK_SKIP_SECTIONS}

    words = []
    skipped = []
    current_section = None

    for line in text.split("\n"):
        heading = match_heading_line(line, found_headings)

        if heading:
            current_section = heading

            if heading.lower() in skip_lookup and heading not in skipped:
                skipped.append(heading)

            # Heading text itself is fixed boilerplate -- nothing to check.
            continue

        if current_section and current_section.lower() in skip_lookup:
            continue

        words.extend(tokenize_for_spelling(line))

    return spell.unknown(words), skipped

def check_guidelines(text):
    results = {}

    # Example guideline: required phrases
    required_phrases = [""]
    missing_phrases = [p for p in required_phrases if p.lower() not in text.lower()]
    results['missing_phrases'] = missing_phrases

    # Example guideline: forbidden words
    forbidden_words = ["error!", "unauthorized"]
    found_forbidden = [w for w in forbidden_words if w.lower() in text.lower()]
    results['forbidden_words'] = found_forbidden

    return results

def check_headings(found_headings):
    """Describe every heading found in the document, in document order.

    Matching against REQUIRED_HEADINGS is case-insensitive, so a heading typed
    "PURPOSE" still counts as the required "Purpose" rather than showing up as
    missing and extra at once. Duplicates are collapsed but counted, so a
    document that repeats "Procedures" three times reports it once with a count.
    """
    required_lookup = {h.lower(): h for h in REQUIRED_HEADINGS}

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
            "text": required_lookup.get(key, text),
            "required": key in required_lookup,
            "count": 1
        }

        by_key[key] = entry
        all_headings.append(entry)

    matched = [h["text"] for h in all_headings if h["required"]]
    extra = [h["text"] for h in all_headings if not h["required"]]
    missing = [h for h in REQUIRED_HEADINGS if h not in matched]

    return {
        "all": all_headings,
        "missing": missing,
        "found": matched,
        "extra": extra,
        "required": REQUIRED_HEADINGS
    }

def format_text_with_headings(text, headings):
    lines = text.split("\n")
    heading_keys = {heading_key(h) for h in headings if heading_key(h)}
    html = ""

    for line in lines:
        # Escape document text -- it is rendered into the results page as HTML.
        safe_line = escape(line)

        if heading_key(line) in heading_keys:
            html += f"<h2>{safe_line}</h2>"
        else:
            html += f"<p>{safe_line}</p>"

    return html


def looks_like_heading(line):
    """Heuristic heading test for formats with no style information."""
    stripped = SECTION_NUMBER.sub("", line.strip()).strip()

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
    required_lookup = {h.lower(): h for h in REQUIRED_HEADINGS}

    found = []
    seen = set()

    for line in text.split("\n"):
        stripped = line.strip()

        if not stripped:
            continue

        # "1. Purpose" and "Purpose" are the same section.
        bare = SECTION_NUMBER.sub("", stripped).strip()
        key = bare.lower()

        if key in required_lookup:
            name = required_lookup[key]
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


@app.route('/', methods=['GET', 'POST'])
def upload_file():

    if request.method == 'POST':

        uploaded_file = request.files.get('document')

        if not uploaded_file:
            return '<h3>No file uploaded</h3>'

        if uploaded_file.filename.endswith('.docx'):

            text, headings = extract_docx_data(uploaded_file.stream)

        elif uploaded_file.filename.endswith('.pdf'):

            reader = PdfReader(uploaded_file.stream)

            text = ""

            for page in reader.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

            headings = find_headings_in_text(text)

        elif uploaded_file.filename.endswith('.txt'):

            text = uploaded_file.stream.read().decode(
                "utf-8",
                errors="ignore"
            )

            headings = find_headings_in_text(text)

        else:

            return "<h3>Unsupported file type</h3>"

        spelling_errors, skipped_sections = check_spelling(text, headings)
        guideline_results = check_guidelines(text)
        heading_results = check_headings(headings)
        formatted_text = format_text_with_headings(text, headings)

        return render_template(
            "results.html",
            filename=uploaded_file.filename,
            all_headings=heading_results["all"],
            found_headings=heading_results["found"],
            missing_headings=heading_results["missing"],
            extra_headings=heading_results["extra"],
            required_headings=heading_results["required"],
            spelling_errors=spelling_errors,
            skipped_sections=skipped_sections,
            missing_phrases=guideline_results["missing_phrases"],
            forbidden_words=guideline_results["forbidden_words"],
            formatted_text=formatted_text
        )

    return render_template("index.html")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
