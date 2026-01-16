# Create sheets for the CMS masterclass

## Installation

```
git clone git@github.com:QuarkNet-HEP/cms-masterclass-sheets.git
cd cms-masterclass-sheets
virtualenv qn
source qn/bin/activate
pip install -r requirements.txt
```

## Usage

This application reads in a template Google Sheet from it creates a new sheet ready for use
in the CMS masterclass.

There are details regarding sheet locations, permissions, and authentications that are not covered here.

```
$ python3 create_sheet2_optimized.py --help

Usage: create_sheet2_optimized.py [OPTIONS]

Options:
  --name TEXT              Add a new spreadsheet with the name: --sheetname
                           <str>  [required]
  --tab <TEXT INTEGER>...  Add a new tab (name and number of datasets): --tab
                           <str> <int> (repeatable)  [required]
  --help                   Show this message and exit.
```

For example:
```
python3 create_sheet2_optimize.py --name "16 Jan, Boston" --tab "Cambridge" 45 --tab "Somerville" 34 --tab "Malden" 10
```