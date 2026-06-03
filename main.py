from flask import Flask, request, render_template
from docx import Document
from PyPDF2 import PdfReader
from spellchecker import SpellChecker
import io

app = Flask(__name__)



def extract_docx_data(file_stream):
    doc = Document(io.BytesIO(file_stream.read()))
    
    full_text = []
    headings = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        full_text.append(text)

        # Detect headings based on style
        if para.style.name.startswith('Heading'):
            headings.append(text)

    return "\n".join(full_text), headings

def check_spelling(text):
    spell = SpellChecker()
    words = text.split()
    misspelled = spell.unknown(words)
    return misspelled

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

    missing = [h for h in REQUIRED_HEADINGS if h not in found_headings]

    return {
        "missing": missing,
        "found": found_headings,
        "required": REQUIRED_HEADINGS
    }
    
def format_text_with_headings(text, headings):
    lines = text.split("\n")
    html = ""

    for line in lines:
        if line in headings:
            html += f"<h2 style='color:blue;'>{line}</h2>"
        else:
            html += f"<p>{line}</p>"

    return html


def find_headings_in_pdf(text):

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

    found = []

    for heading in REQUIRED_HEADINGS:

        if heading.lower() in text.lower():
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

            headings = find_headings_in_pdf(text)

        elif uploaded_file.filename.endswith('.txt'):

            text = uploaded_file.stream.read().decode(
                "utf-8",
                errors="ignore"
            )

            headings = []

        else:

            return "<h3>Unsupported file type</h3>"

        spelling_errors = check_spelling(text)
        guideline_results = check_guidelines(text)
        heading_results = check_headings(headings)
        formatted_text = format_text_with_headings(text, headings)

        return render_template(
            "results.html",
            filename=uploaded_file.filename,
            found_headings=heading_results["found"],
            missing_headings=heading_results["missing"],
            required_headings=heading_results["required"],
            spelling_errors=spelling_errors,
            missing_phrases=guideline_results["missing_phrases"],
            forbidden_words=guideline_results["forbidden_words"],
            formatted_text=formatted_text
        )

    return render_template("index.html")

if __name__ == '__main__':
    app.run(debug=True)