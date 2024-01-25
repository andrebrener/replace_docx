# Generate Docx files programatically

Using data provided on a Google spreadsheet, generate docx files replacing the data.

To replace variables, set them with the name `${variable_name}` on the docx model.

This repo uses [python-docx-replace](https://pypi.org/project/python-docx-replace/) and [python-docx](https://python-docx.readthedocs.io/en/latest/index.html)

For more info check their documentations.

## Getting Started

### 1. Clone Repo

`git clone https://github.com/andrebrener/replace_docx.git`

### 2. Install Packages Required

Go in the directory of the repo and run:
`pip install -r requirements.txt`

### 3. Create a `.env` file

`touch .env`

### 3. Set the following variables in the `.env` file

- `SHEET_ID` = Google spreadsheet id
- `SHEET_NAME` = Name of the sheet inside google sheets
- `FILE_NAME_MODEL` = Name of the docx model to replace

### 4. Get Files

- Run [docx_replace.py](https://github.com/andrebrener/replace_docx/blob/main/docx_replace.py).
- When the script finishes, files will be created under the `results` directory
