#!/usr/bin/env python3
import docx
import re
from docx.enum.text import WD_COLOR_INDEX
import os

script_path = os.path.dirname(os.path.abspath( __file__ ))
file_path = '/Desktop/autodate/resources/'

doc = docx.Document(script_path + file_path + 'fiche_histoire.docx')

regex = r"(\d{4}\s?[:])|(\d{4}\s*[=]\s?[>])|(\d{4}\s?[-]\s?\d{4})|([A]\d{2})"

for p in doc.paragraphs:
    for run in p.runs:
        date1 = re.findall(regex, run.text)
        if date1:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW

doc.save(script_path + file_path + 'restyled.docx')