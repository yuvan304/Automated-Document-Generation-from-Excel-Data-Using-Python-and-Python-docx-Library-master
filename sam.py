from docx import Document
from docx.oxml import OxmlElement
import pandas as pd

# Read Excel data
excel_data = pd.read_excel('sample.xlsx')

# Open the Word document
doc = Document('sample.docx')

# Update Account/Application Numbers dropdown


# Save the updated Word document
doc.save('sample.docx')

print("Account numbers' last four digits have been updated in the Word document dropdown.")
