# ToolBox

This repository is intended to host many versatile scripts to automate common tasks.

The scripts are grouped by language/context and with the same principle will be shown here. 

## Python

Under the _python_ directory there will be all the Python scripts with a contextual 
`requirements.txt` file to set up the `.venv`.

### xls2json

This script is used to transpose an _xls_ or _xlsx_ workbook sheet to a json list of records.

The script accepts a list of field names that will be used to compose each json object 
corresponding to a row. One of such fields must be used as a unique key to prevent record
duplication.
