# Certificate-Generator
📄 Certificate Generator

This project automates the generation of personalized Certificates of Achievement using:

Excel data (Name + Content)

A DOCX certificate template

Python script to generate DOCX & PDF certificates

It reads each entry from the Excel file, fills the Word template with the provided information, saves the output as DOCX, and then converts all certificates into PDF.

🚀 Features

✔ Reads data from results.xlsx

✔ Automatically fills placeholders in template.docx

✔ Generates individual DOCX certificates

✔ Converts all DOCX files to PDF format

✔ Automatically creates an output folder

✔ Filename-safe generation for each certificate

📂 Project Structure
project/
│── generate_latters.py
│── results.xlsx
│── template.docx
│── output/         (auto-generated)
│     ├── Name1.docx
│     ├── Name1.pdf
│     ├── ...

🛠️ Requirements

Install required dependencies:

pip install pandas docxtpl docx2pdf

Additional Requirements for PDF Conversion

Windows: Requires Microsoft Word

Linux/macOS: Requires LibreOffice

If PDF conversion fails, DOCX files will still be generated.

▶️ How to Run

Place your Excel file as results.xlsx

Ensure your DOCX template is named template.docx

Open terminal in project folder

Run the script:

python generate_latters.py

📄 Excel Format (results.xlsx)

The Excel file must contain the following columns:

Name	Content
Student Name	Description / Achievement text

Example:

Name	Content
Samiksha	Successfully completed training
Priya	Outstanding performance in ML
📝 Template Format (template.docx)

Your Word template must include the placeholders:

{{name}}
{{content}}


These will be automatically replaced for each row.

🧠 How It Works
1. Read Excel File
df = pd.read_excel(excel_file)

2. Render Template for Each Entry
doc.render({"name": name, "content": content})

3. Save DOCX

Files are saved in /output with safe filenames.

4. Convert All DOCX → PDF
convert(output_folder)

📢 Troubleshooting
❌ ModuleNotFoundError

Install missing packages:

pip install docxtpl pandas docx2pdf

❌ PDF Conversion Fails

Ensure MS Word/LibreOffice is installed

Run as Administrator (Windows)

Try deleting the output folder and running again
