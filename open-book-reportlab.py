#!/usr/bin/env python3

#
# Copyright Colin Turner 2014-2026
#
# Free and Open Source Software under GPL v3
#
"""Produce personalised guide sheets for semi-open-book examinations without LaTeX.

An alternative to open-book.py that needs only Python packages:

* by default the whole sheet is drawn with reportlab, using the wording below;
* with --template-pdf, an existing one-page PDF (for example a sheet designed
  in Word and saved as PDF) is used instead, and only the student number (and
  optional barcode) is stamped onto it.

Requires:  pip install reportlab pypdf
"""

import sys
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from open_book_common import SheetError, finish_arguments, make_parser, run

try:
    from reportlab.graphics.barcode.code39 import Standard39
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Paragraph
except ImportError:
    sys.exit('This script needs reportlab: pip install reportlab pypdf')

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None   # Only needed for --template-pdf

# The wording of the sheet when no --template-pdf is given. {module} is
# replaced by the module code. <b>...</b> gives bold text.
TITLE = '<b>{module} Examination Guide Sheet</b>'
INSTRUCTIONS = (
    'This sheet, and its contents that you have added, can be brought into the '
    'examination for {module}. The contents <b>must</b> be compiled by yourself, be '
    'handwritten, and be original (i.e. <b>NOT</b> photocopied or similar). You may use '
    'the reverse side. You may retain a copy you have made before the examination but '
    'the original must be handed in with your examination scripts at the end of your '
    'examination.')
DECLARATION = (
    'Please read the following declaration and tick the box to indicate you agree:<br/>'
    'I declare this sheet to have been compiled by myself and not by another, and that '
    'the student number above is mine.')

MARGIN = 20 * mm
FONT, BOLD_FONT, NUMBER_FONT = 'Helvetica', 'Helvetica-Bold', 'Courier'
BODY = ParagraphStyle('body', fontName=FONT, fontSize=12, leading=15, alignment=TA_CENTER)


def draw_paragraph(pdf, markup, top, width):
    """Draw centred paragraph text with its top at y=top; return the y of its bottom."""
    para = Paragraph(markup, BODY)
    _, height = para.wrap(width, A4[1])
    para.drawOn(pdf, MARGIN, top - height)
    return top - height


def draw_number(pdf, args, student_number, right, top):
    """Draw the student number (with optional barcode above it), right-aligned at x=right below y=top."""
    if args.barcode:
        # checksum=0 matches the plain Code 39 the LaTeX version produces
        barcode = Standard39(student_number, barWidth=0.4 * mm, barHeight=10 * mm,
                             ratio=2.5, checksum=0, quiet=0)
        top -= barcode.height
        barcode.drawOn(pdf, right - barcode.width, top)
        top -= 2 * mm
    pdf.setFont(NUMBER_FONT, 14)
    pdf.drawRightString(right, top - 14, student_number)


def draw_full_sheet(pdf, args, student_number):
    """Draw the complete guide sheet, laid out like the LaTeX template."""
    page_width, page_height = A4
    width = page_width - 2 * MARGIN
    module = escape(args.module_code)

    y = draw_paragraph(pdf, TITLE.format(module=module), page_height - MARGIN, width)
    y = draw_paragraph(pdf, INSTRUCTIONS.format(module=module), y - 4 * mm, width)
    draw_number(pdf, args, student_number, page_width - MARGIN, y - 4 * mm)

    # The declaration and tick box sit at the bottom of the page
    box = 3 * mm
    pdf.rect((page_width - box) / 2, MARGIN, box, box)
    declaration = Paragraph(DECLARATION.format(module=module), BODY)
    declaration.wrap(width, page_height)
    declaration.drawOn(pdf, MARGIN, MARGIN + box + 3 * mm)


def create_pdf(args, student_number, destination):
    """Make one student's guide sheet and save it to destination."""
    try:
        make_sheet(args, student_number, destination)
    except Exception as error:
        raise SheetError('could not make the sheet for {}: {}'.format(student_number, error))


def make_sheet(args, student_number, destination):
    """Draw the sheet, or stamp the template PDF, and write it to destination."""
    if args.template_pdf is None:
        pdf = canvas.Canvas(str(destination), pagesize=A4)
        pdf.setTitle('{} Examination Guide Sheet'.format(args.module_code))
        draw_full_sheet(pdf, args, student_number)
        pdf.save()
        return

    # Stamp the number onto the first page of the template PDF
    page = PdfReader(args.template_pdf).pages[0]
    width, height = float(page.mediabox.width), float(page.mediabox.height)
    overlay = BytesIO()
    pdf = canvas.Canvas(overlay, pagesize=(width, height))
    draw_number(pdf, args, student_number, width - args.stamp_right * mm, height - args.stamp_top * mm)
    pdf.save()

    writer = PdfWriter(clone_from=args.template_pdf)
    writer.pages[0].merge_page(PdfReader(overlay).pages[0])
    with open(destination, 'wb') as f:
        writer.write(f)


def main():
    parser = make_parser('Produce, and optionally email, guide sheets for a semi-open-book '
                         'exam without LaTeX (uses reportlab)')
    parser.add_argument('--template-pdf', type=Path,
                        help='a one-page PDF (e.g. designed in Word) to stamp the student number '
                             'onto, instead of drawing the whole sheet')
    parser.add_argument('--stamp-top', type=float, default=20,
                        help='with --template-pdf: distance of the number from the top of the page, '
                             'in mm (default 20)')
    parser.add_argument('--stamp-right', type=float, default=20,
                        help='with --template-pdf: distance of the number from the right of the page, '
                             'in mm (default 20)')
    args = finish_arguments(parser.parse_args())

    if args.template_pdf is not None:
        if PdfReader is None:
            sys.exit('--template-pdf needs pypdf: pip install pypdf')
        if not args.template_pdf.exists():
            sys.exit('Cannot find the template {}'.format(args.template_pdf))

    run(args, create_pdf)


if __name__ == '__main__':
    main()
