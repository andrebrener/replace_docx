from docx import Document
import pandas as pd 
from python_docx_replace import docx_replace
import os
from dotenv import load_dotenv

load_dotenv()

sheet_id = os.environ['SHEET_ID']
sheet_name = os.environ['SHEET_NAME']
file_name_model = os.environ['FILE_NAME_MODEL']


## Data to complete

url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}'

df = pd.read_csv(url)

## Strip all values

for c in df.columns:
    df[c] = df[c].astype(str).str.strip()

## Make directory
    
dir = 'results'

os.makedirs(dir, exist_ok=True)

for r in df.iterrows():

    row = r[1]

    contrato_number = row['contrato']

    if contrato_number != '1':
        continue

    print(f"Replacing {contrato_number}")

    ## Open docx

    document = Document(f'{file_name_model}.docx')

    docx_replace(document, **row)

    # do whatever you want after that, usually save the document
    document.save(os.path.join(dir, f"{contrato_number}_replaced_contrato.docx"))






