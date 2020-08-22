# xlsx2docx-byColumns

This project hosts the python3 tool `xlsx2docx-byColumns.py` and after build a derived MS-Windows executable `xlsx2docx-byColumns.exe`

## Workflow

The Python3 script `xlsx2docx-byColumns.py` generates out of an xlsx database (MS-Excel file) the resulting docx files (MS-Word).
An TOML formatted configuration file `xlsx2docx-byColumns_Configuration.toml` provides information to control the input and output data, e.g. file path, used rows and columns.
The script uses the `xlsx2docx-byColumns_Template.docx` to define styles and to replace or add paragraphs controlled by configurable columns in the MS-Excel database `xlsx2docx-byColumns_SourceData_Example.xlsx`.

## Architecture

The source data are in an xlsx sheet (MS-Excel). This database hosts all data that will be transformed into generated docx files (MS-World).

1. SourceDataExample.xlsx

From this database multiple resulting MS-Word files will be derived.

1. generatedDataFile1_yyyy-mm-dd_hhmmss.docx
2. generatedDataFile2_yyyy-mm-dd_hhmmss.docx
3. generatedDataFile3_yyyy-mm-dd_hhmmss.docx

## Target platform

The python script can be used wherever the needed modules in requirements.txt are available.

In addition an standalone executable is intended to be used for the MS-Windows platform.
As it is assumed that the main users are working with Windows.

## Sub-directories

1. src
2. test
3. build

### src

This source folder contains the script to generate the MS-Word results and a configuration file to run the examples in the test folder.

1. Python3 script: `xlsx2docx-byColumns.py`
2. Example TOML configuration file: `xlsx2docx-byColumns_Configuration.toml`

### test

This test folder contains test and example data:

1. Example XLSX database as input:  `xlsx2docx-byColumns_SourceData_Example.xlsx`
2. Example DOCX template:           `xlsx2docx-byColumns_Template_Example.xlsx`

### build

This build folder will be generated and contains several results.

## Development Environment

This project is developed under MS-Windows using VSCode and was tested partly with WSL.

Install dependencies for Python3 (on the cmd.exe or powershell.exe):
    py -3 -m pip install -r requirements.txt

Run the script on MS-Windows via a small wrapper (to support the projects folder structure) on the cmd.exe or powershell.exe:
    .\xlsx2docx-byColumns_run_script.bat

Install the dependency to build the additional standalone executable:
    py -3 -m pip install pyinstaller

Build the additional standalone executable:
    pyinstaller src/xlsx2docx-byColumns.py --onefile --distpath ./build/dist --specpath ./build

Run the executable on MS-Windows via a small wrapper (to support the projects folder structure) on the cmd.exe or powershell.exe:
    .\xlsx2docx-byColumns_run_exe.bat

## Additional information sources

[1] Inspired by "Automate the boring stuff with python" <https://automatetheboringstuff.com/2e/chapter13/>  <https://automatetheboringstuff.com/2e/chapter15/>
