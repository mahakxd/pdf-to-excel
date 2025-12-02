A small tool that reads information from a PDF and converts it into a structured Excel and CSV format using Gemini 2.5 Flash.
The goal of the project is to take unorganized text (like a profile, resume, or narrative description) and turn it into clean tabular data.

Overview:

>The app extracts raw text from a PDF, sends it to Gemini for parsing, and then formats the response into:

>A refined Excel file (with column sizes, borders, header formatting)

>A CSV file

>A live preview inside the UI

>This helps in cases where data needs to be standardized, indexed, or used for analytics.

Features:

>Converts unstructured PDF text into key–value–comment rows

>Shows the extracted data on screen (no download needed)

>Saves both CSV and Excel files locally inside an output/ folder

>Excel file is formatted (column width, headers, borders)

>Simple Gradio interface for testing

Tech Used:

>Python

>pdfplumber (PDF text reading)

>google-genai (Gemini API)

>pandas

>openpyxl (Excel styling)

>Gradio (UI)

How to Run:

Install dependencies:

pip install -r requirements.txt


Add your Gemini API key inside app.py:

genai_client = genai.Client(api_key="YOUR_API_KEY")


Start the app:

python app.py


Upload a PDF → view the extracted table → download Excel if needed.

Project Structure
app.py
requirements.txt
output/
    structured_output.xlsx
    structured_output.csv

Notes

The model is responsible for deciding which parts are keys, values, and comments.

The app does not use OCR; it expects normal machine-readable PDFs.

Output files get overwritten each time you process a new PDF.

Possible Improvements

Support for multi-page grouping

Better error handling around inconsistent model responses

Deployment as a web app

Optional history of processed files

Dark mode UI

Why I Built It

The assignment required an automated way to take raw textual data and turn it into something structured.
This solution focuses on reliability, readability of the output, and ease of use for someone who may not be technical.
