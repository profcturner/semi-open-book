#
# Copyright Colin Turner 2014-2026
#
# Free and Open Source Software under GPL v3
#
"""Code shared by the semi-open-book scripts.

Reading the student CSV file, command line options, prompting, and emailing
are the same whichever way the PDF is made. Each script supplies a function
that makes one student's PDF and calls run().
"""

import argparse
import csv
import getpass
import os
import re
import smtplib
import sys
from email.message import EmailMessage
from pathlib import Path

DEFAULT_EMAIL_BODY = ('Dear Student\n'
                      'Please find enclosed your guide sheet template for the exam. '
                      'Read the following email carefully.\n')


class SheetError(Exception):
    """Raised when a guide sheet cannot be made for a student."""


def make_parser(description):
    """Return an ArgumentParser with the options common to all the scripts."""
    parser = argparse.ArgumentParser(
        description=description,
        epilog='Note that column count arguments start from zero.')

    parser.add_argument('-b', '--batch-mode', action='store_true',
                        help='run automatically with values given, without prompting')
    parser.add_argument('--interactive-mode', action='store_true', default=True,
                        help='prompt the user for details (default)')
    parser.add_argument('-i', '--input-file', default='students.csv',
                        help='the name of the input CSV file with one row per student')
    parser.add_argument('--csv-encoding', default='utf-8-sig',
                        help='the text encoding of the CSV file (default utf-8-sig; '
                             'try cp1252 for older Excel files)')
    parser.add_argument('-m', '--module-code', default='MODULE',
                        help='the module code printed on the sheet, e.g. EEE122')
    parser.add_argument('-o', '--output-dir', type=Path, default=Path('sheets'),
                        help='the folder in which to save the PDFs (default sheets)')
    parser.add_argument('--barcode', action='store_true',
                        help='also print the student number as a Code 39 barcode')
    parser.add_argument('-sidc', '--student-id-column', type=int, default=1,
                        help='the column containing the student id (default 1)')
    parser.add_argument('-snc', '--student-name-column', type=int, default=2,
                        help='the column containing the student name, used only for '
                             'progress messages (default 2)')
    parser.add_argument('-sec', '--student-email-column', type=int, default=9,
                        help='the column containing the student email (default 9)')
    parser.add_argument('-sidregexp', '--student-id-regexp', default='B[0-9]+',
                        help='a regular expression the whole student ID must match (default B[0-9]+)')
    parser.add_argument('--smtp-server', default='localhost',
                        help='the address of an SMTP server')
    parser.add_argument('--smtp-port', type=int, default=25,
                        help='the SMTP port (default 25; usually 587 with --smtp-starttls)')
    parser.add_argument('--smtp-starttls', action='store_true',
                        help='encrypt the SMTP connection with STARTTLS')
    parser.add_argument('--smtp-user', default='',
                        help='log in to the SMTP server as this user; the password is prompted for, '
                             'or taken from the OPEN_BOOK_SMTP_PASSWORD environment variable')
    parser.add_argument('--email-subject', default='IMPORTANT: Your semi-open-book Guide Sheet',
                        help='the subject of emails that are sent')
    parser.add_argument('--email-sender', default='noreply@nowhere.org',
                        help='the sender address from which to send emails')
    parser.add_argument('--email-body-file', type=Path,
                        help='a text file containing the body of the email')
    parser.add_argument('-t', '--test-only', '--no-email', action='store_true', dest='test_only',
                        help='only create the PDFs, do not send any emails')
    return parser


def prompt(question, current, convert=str):
    """Ask the user for a value, keeping the current one if they just press Enter."""
    answer = input('{}? default=[{}] :'.format(question, current)).strip()
    return convert(answer) if answer else current


def finish_arguments(args):
    """Prompt for overrides in interactive mode, then fill in derived values."""
    if args.interactive_mode and not args.batch_mode:
        args.input_file = prompt('CSV filename', args.input_file)
        args.module_code = prompt('Module code', args.module_code)
        args.student_id_column = prompt('Student ID Column', args.student_id_column, int)
        args.student_name_column = prompt('Student Name Column', args.student_name_column, int)
        args.student_email_column = prompt('Student Email Column', args.student_email_column, int)
        args.student_id_regexp = prompt('Student ID Regular Expression', args.student_id_regexp)
        if not args.test_only:
            args.smtp_server = prompt('SMTP Server', args.smtp_server)
            args.smtp_user = prompt('SMTP username (blank for none)', args.smtp_user)
            args.email_subject = prompt('Email subject', args.email_subject)
            args.email_sender = prompt('Email sender address', args.email_sender)

    args.email_body = (args.email_body_file.read_text(encoding='utf-8')
                       if args.email_body_file else DEFAULT_EMAIL_BODY)
    return args


def read_students(args):
    """Yield (student_number, student_name, student_email) for each CSV row with a valid student number."""
    needed = max(args.student_id_column, args.student_name_column, args.student_email_column)
    with open(args.input_file, 'r', encoding=args.csv_encoding, newline='') as f:
        for line_number, row in enumerate(csv.reader(f, dialect='excel'), start=1):
            if len(row) <= needed:
                print('  Skipping line {}: too few columns'.format(line_number))
                continue
            student_number = row[args.student_id_column].strip()
            if not re.fullmatch(args.student_id_regexp, student_number):
                print('  Skipping line {}: no valid student number'.format(line_number))
                continue
            yield (student_number, row[args.student_name_column].strip(),
                   row[args.student_email_column].strip())


def build_email(args, student_email, pdf_path):
    """Create the email for one student with their PDF attached."""
    msg = EmailMessage()
    msg['Subject'] = args.email_subject
    msg['From'] = args.email_sender
    msg['To'] = student_email
    msg.set_content(args.email_body)
    msg.add_attachment(pdf_path.read_bytes(), maintype='application', subtype='pdf',
                       filename=pdf_path.name)
    return msg


def connect_smtp(args):
    """Open a connection to the SMTP server, using STARTTLS and logging in if asked."""
    server = smtplib.SMTP(args.smtp_server, args.smtp_port)
    if args.smtp_starttls:
        server.starttls()
    if args.smtp_user:
        password = os.environ.get('OPEN_BOOK_SMTP_PASSWORD') or \
            getpass.getpass('SMTP password for {}: '.format(args.smtp_user))
        server.login(args.smtp_user, password)
    return server


def run(args, make_pdf):
    """Make (and optionally email) a sheet for every student in the CSV file.

    make_pdf(args, student_number, destination) must write the PDF to the
    destination path, raising SheetError if it cannot.
    """
    args.output_dir.mkdir(parents=True, exist_ok=True)

    print('Starting open-book...')
    smtp = None if args.test_only else connect_smtp(args)
    done, failed = 0, []
    try:
        for student_number, student_name, student_email in read_students(args):
            print('  Processing:', student_number, ':', student_name, ':', student_email)
            destination = args.output_dir / '{}-{}.pdf'.format(args.module_code, student_number)
            try:
                # Remove any old copy so a failure can never leave a stale sheet to be sent
                destination.unlink(missing_ok=True)
                make_pdf(args, student_number, destination)
                if smtp:
                    smtp.send_message(build_email(args, student_email, destination))
                done += 1
            except (SheetError, smtplib.SMTPException) as error:
                print('  FAILED:', error, file=sys.stderr)
                failed.append(student_number)
    finally:
        if smtp:
            smtp.quit()

    print('Created {} sheet(s) in {}{}'.format(
        done, args.output_dir, '' if args.test_only else ' and emailed them'))
    if failed:
        print('Failed for: {}'.format(', '.join(failed)), file=sys.stderr)
        sys.exit(1)
