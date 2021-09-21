"""
Generate PDF Tables from Excel Files

Author: Jan Thilo Savary

Codefragments from https://pbpython.com/pdf-reports.html by Chris Moffitt
"""

import argparse

import pandas as pd
import numpy as np
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from pathlib import Path
from sys import stderr


def pagenumber_type(x):
    x = int(x)
    if x < 1:
        raise argparse.ArgumentTypeError("Pagenumber must be greater than 0")
    return x


def main(args):
    df = pd.read_excel(args.infile.name)


    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template("templates/report.html")

    template_vars = {"table": df.to_html(index=False),
                     "align": args.align,
                     "pagenumber": args.pagenumber,
                     'pagenumberalign': args.pagenumberalign}
    if args.title is not None:
        template_vars['title'] = args.title
    if args.fontsize is not None:
        template_vars['fontsize'] = args.fontsize
    if args.landscape:
        template_vars['landscape'] = True
    if args.disablepagenumber:
        template_vars['disablepagenumber'] = True

    html_out = template.render(template_vars)

    if args.html:
        with open(Path(args.outfile.name).with_suffix('.html'), "w", encoding="utf-8") as file:
            file.write(html_out)

    HTML(string=html_out).write_pdf(args.outfile.name)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate PDF Table.')
    parser.add_argument('-i', '--infile', type=argparse.FileType('r'),
                        help="report source file in Excel (Default: data/data.xlsx)", default='data/data.xlsx')
    parser.add_argument('-o', '--outfile', type=argparse.FileType('w'),
                        help="output file in PDF (Default: report.pdf)", default='report.pdf')
    parser.add_argument('--html', help="save html file", action='store_true')
    parser.add_argument('-t', '--title', help="title displayed in document")
    parser.add_argument('-f', '--fontsize', help="font size (must be valid css)")
    parser.add_argument('--align', choices=['left', 'center', 'right'],
                        help="alignment in cells ('left', 'center', 'right')", default='left')
    parser.add_argument('--landscape', action='store_true', help="enable landscape")
    parser.add_argument('-p', '--pagenumber', help="start pagenumber", type=pagenumber_type, default=1)
    parser.add_argument('--disablepagenumber', help="disable pagenumber", action='store_true')
    parser.add_argument('--pagenumberalign', choices=['left', 'center', 'right'],
                        help="alignment of pagenumber ('left', 'center', 'right')", default='right')
    args = parser.parse_args()
    main(args)