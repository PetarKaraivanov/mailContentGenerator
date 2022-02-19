Extract the information from the CSV, replace the placeholders in the docx template and generate pdfs.
CSV: first line holds the placeholder names. Those exact names are found in the docx template.
DOCX Template: a word document containing the placeholders. Note that currently all images will disappear
after the replacement(problem of the python-docx) Only images in the header and the footer will remain.
INSTALLATION: "pip3 install -r requirements.py"
RUN: "python main.py"
