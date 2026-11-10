#!/usr/bin/env python3

#
# Copyright Colin Turner 2014-2026
#
# Free and Open Source Software under GPL v3
#
"""Produce personalised guide sheets for semi-open-book examinations using LaTeX.

Reads a CSV file with one row per student, builds a PDF guide sheet for each
from a LaTeX template with pdflatex, saves them in an output folder and,
optionally, emails each student their sheet.
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from open_book_common import SheetError, finish_arguments, make_parser, run

SCRIPT_DIR = Path(__file__).resolve().parent

# Characters that need escaping in LaTeX
LATEX_SPECIALS = {
    '\\': r'\textbackslash{}', '{': r'\{', '}': r'\}', '$': r'\$', '&': r'\&',
    '#': r'\#', '^': r'\^{}', '_': r'\_', '~': r'\~{}', '%': r'\%',
}


def latex_escape(text):
    """Escape text so that it can be safely placed in a LaTeX document."""
    return ''.join(LATEX_SPECIALS.get(c, c) for c in text)


def write_settings(directory, args, student_number):
    """Write the per-student settings file that the template inputs.

    It defines \\ModuleCode and \\StudentNumber, and turns the barcode on or off.
    """
    lines = [
        r'\newcommand{\ModuleCode}{' + latex_escape(args.module_code) + '}',
        r'\newcommand{\StudentNumber}{' + latex_escape(student_number) + '}',
        r'\ShowBarcodetrue' if args.barcode else r'\ShowBarcodefalse',
    ]
    (directory / 'open-book-settings.tex').write_text('\n'.join(lines) + '\n', encoding='utf-8')


def create_pdf(args, student_number, destination):
    """Build one student's guide sheet with pdflatex and save it to destination.

    Each build happens in a fresh temporary folder, so nothing from one
    student's build can leak into the next.
    """
    with tempfile.TemporaryDirectory() as tmp:
        build = Path(tmp)
        shutil.copy(args.template, build / 'open-book.tex')
        write_settings(build, args, student_number)

        result = subprocess.run(
            [args.pdflatex, '-interaction=nonstopmode', '-halt-on-error', 'open-book.tex'],
            cwd=build, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, errors='replace')
        pdf = build / 'open-book.pdf'
        if result.returncode != 0 or not pdf.exists():
            # The useful part of a LaTeX error is near the end of the log
            tail = '\n'.join(result.stdout.splitlines()[-15:])
            raise SheetError('pdflatex failed for {}:\n{}'.format(student_number, tail))
        shutil.move(str(pdf), destination)


def main():
    parser = make_parser('Produce, and optionally email, guide sheets for a semi-open-book '
                         'exam using LaTeX')
    parser.add_argument('--template', type=Path, default=SCRIPT_DIR / 'open-book.tex',
                        help='the LaTeX template to use (default open-book.tex next to this script)')
    parser.add_argument('--pdflatex', default='pdflatex',
                        help='the pdflatex command to run (default pdflatex)')
    args = finish_arguments(parser.parse_args())

    if not shutil.which(args.pdflatex):
        sys.exit('Cannot find {}: please install a TeX distribution '
                 '(e.g. MiKTeX or TeX Live).'.format(args.pdflatex))
    if not args.template.exists():
        sys.exit('Cannot find the template {}'.format(args.template))

    run(args, create_pdf)


if __name__ == '__main__':
    main()
