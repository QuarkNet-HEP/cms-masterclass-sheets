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

In addition, one can create a Combination sheet which collects information from the "Results" tab over multiple sheets:

```
$ python3 create_combination.py --help

Usage: create_combination.py [OPTIONS]

Options:
  --name TEXT   Add a new spreadsheet with the name: --sheetname <str>
                [required]
  --sheet TEXT  : Specify sheets to summarize by url --sheet <str>  [required]
  --help        Show this message and exit.
```

## Developing

### Linting

[Ruff](https://docs.astral.sh/ruff/) is used for linting and formatting. It is installed as part of `requirements.txt`.

To run ruff locally:
```
# check for errors
ruff check create_combination.py create_sheet2_optimized.py

# check and auto-fix
ruff check --fix create_combination.py create_sheet2_optimized.py

# format
ruff format create_combination.py create_sheet2_optimized.py
```

### Pre-commit

[pre-commit](https://pre-commit.com/) is used to run ruff automatically before each commit. It is installed as part of `requirements.txt`.

After installing dependencies, enable the hook with:
```
pre-commit install
```

Ruff will then run automatically on staged files whenever you `git commit`. To run it manually:
```
pre-commit run --all-files
```
