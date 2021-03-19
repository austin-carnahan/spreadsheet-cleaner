# USDR Spreadsheet Cleaner

## Project Overview:
Automate the data ingestion, cleaning, collating and aggregation pipeline for COVID-19 data

Contains spreadsheet deduplication scripts as well as a main gui-based spreadsheet cleaner. Spreadsheet Cleaner gui uses rules spreadsheets
to track valid spreadsheet values and automatically applies provided data transformations, so you can effectively note a change once and
not have to repeat it. Ideal for screening data pipelines that regularly include spreadsheets.

Project files are on google drive

## Creating a develpment environment

`pip install -r requirements.txt`

run the program:

`python main.py`


## Building a Windows executable

Tested using pyinstaller on windows. This will create `build` and `dist` folders.
`dist` contains a windows executable in `DataTool/DataTool.exe` as well as neccessary configuration files and data directories for program execution

Navigate to the root project folder and run:

`pyinstaller --clean --noconfirm DataTool.spec`

`DataTool.spec` specifies the name of the executable to be created and as well as which directories and files need to be copied for runtime execution.
these include:
- tutorial/
- templates/
- config.ini

The `DataTool` folder can be copied from `dist`, compressed, and distributed as needed. Run `DataTool.exe` to launch the program.
All files and folders within the `DataTool` directory are required for the program to function.

## Publishing a release

Iterate version number in DataTool.spec, build a windows executable, and publish using Github releases.



