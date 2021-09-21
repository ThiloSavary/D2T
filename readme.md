# D2T

D2T is a small, hacky way to export Excel Worksheets into PDF Tables.

## Installation

- First check whether all requirements are met.
- (Setup a virtual enviroment)
- Install the required python packages listed in [requirements.txt](requirements.txt).
    <br> ``pip install -r requirements.txt``

## Usage

```
python main.py
```

Arguments

- i, --infile <br> Input file path **(data/data.xlsx)**
- o, --outfile <br> Output file path **(report.pdf)**
- t, --title <br> Title in Document
- f, --fontsize <br> Fontsize (must be valid css)
- --html <br> Save corresponding html file
- --align <br> Cell alignment *{**left**, right, center}*
- --landscape <br> Enable landscape
- -p, --pagenumber <br> starting pagenumber **(1)**
- --disablepagenumber, <br> disable pagenumber
- --pagenumberalign <br> pagenumber alignment *{left, **right**, center}*

**Example:**

``python main.py --html --align left -f "0.5em" -p 3 --landscape --pagenumberalign center``

## Requirements 

- https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#installation 
    <br> Make sure, you have the requirements for WasyPrint installed. On Windows you have to install GTK.
    Check if the PATH Variable is correct after installation.

## Sources

- Codefragments from https://pbpython.com/pdf-reports.html by Chris Moffitt