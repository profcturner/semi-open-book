#!/usr/bin/env python

#
# Copyright Colin Turner 2014-2017
#
# Free and Open Source Software under GPL v3
#
import argparse

import csv
import re
import subprocess
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def process_student(args, row):
    """Takes a line from the CSV file, you will likely need to edit aspects of this."""
    'kicks of the processing of a single student'
    student_number = row[args.student_id_column]
    student_name = row[args.student_name_column]
    student_email = row[args.student_email_column]
    print('  Processing:', student_name , ':', student_email)
    create_latex_inserts(student_number, student_name)
    create_pdf()
    send_email(args, student_name, student_email)


def create_latex_inserts(student_number, student_name):
    """Write LaTeX inserts for the barcode and student name

    For each student this will create two tiny LaTeX files:

     * open-book-insert-barcode.tex which contains the LaTeX code for a barcode representing the student number
     * open-book-insert-name.tex which will contain simply the student's name

    These files can be included/inputted from open-book.tex as desired to personalise that document

    student_number is the ID in the students record system for the student
    student_name is the name of the student"""

    # Open a tiny LaTeX file to put this in
    file = open('open-book-insert-barcode.tex', 'w')

    # All the file contains is LaTeX to code to create the bar code
    string = '\psbarcode{' + student_number + '}{includetext height=0.25}{code39}'
    file.write(string)
    file.close()

    # The same exercise for the second file to contain the student name
    file = open('open-book-insert-name.tex', 'w')
    string = student_name
    file.write(string)
    file.close()


def create_pdf():
    """Calls LaTeX and dvipdf to create the personalised PDF with inserts from create_latex_inserts()"""

    # Suppress stdout, but we leave stderr enabled.
    subprocess.call("latex open-book", stdout=subprocess.DEVNULL, shell=True)
    subprocess.call("dvipdf open-book", stdout=subprocess.DEVNULL, shell=True)


def send_email(args, student_name, student_email):
    """Emails a single student with the generated PDF."""
    #TODO: Might be useful to improve the to address
    #TODO: Allow subject to be tailored.

    subject = args.email_subject
    from_address = args.email_sender
    # to_address = student_name + ' <' + student_email + '>'
    to_address = student_email

    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_address
    msg['To'] = to_address

    text = 'Dear Student\nPlease find enclosed your guide sheet template for the exam. Read the following email carefully.\n'
    part1 = MIMEText(text, 'plain')
    msg.attach(part1)

    # Open the files in binary mode.  Let the MIMEImage class automatically
    # guess the specific image type.
    fp = open('open-book.pdf', 'rb')
    img = MIMEApplication(fp.read(), 'pdf')
    fp.close()

    msg.attach(img)

    # Send the email via our own SMTP server, if we are not testing.
    if not args.test_only:
        s = smtplib.SMTP(args.smtp_server)
        s.sendmail(from_address, to_address, msg.as_string())
        s.quit()


def override_arguments(args):
    """If necessary, prompt for arguments and override them

    Takes, as input, args from an ArgumentParser and returns the same after processing or overrides.
    """

    # If the user enabled batch mode, we disable interactive mode
    if args.batch_mode:
        args.interactive_mode = False

    if args.interactive_mode:
        override = input("CSV filename? default=[{}] :".format(args.input_file))
        if len(override):
            args.input_file = override

        override = input("Student ID Column? default=[{}] :".format(args.student_id_column))
        if len(override):
            args.student_id_column = int(override)

        override = input("Student Name Column? default=[{}] :".format(args.student_name_column))
        if len(override):
            args.student_name_column = int(override)

        override = input("Student Email Column? default=[{}] :".format(args.student_email_column))
        if len(override):
            args.student_email_column = int(override)

        override = input("Student ID Regular Expression? default=[{}] :".format(args.student_id_regexp))
        if len(override):
            args.student_id_regexp = override

        override = input("SMTP Server? default=[{}] :".format(args.smtp_server))
        if len(override):
            args.smtp_server = override

        override = input("Email subject? default=[{}] :".format(args.email_subject))
        if len(override):
            args.email_subject = override

        override = input("Email sender address? default=[{}] :".format(args.email_sender))
        if len(override):
            args.email_sender = override

    return(args)


def parse_arguments():
    """Get all the command line arguments for the file and return the args from an ArgumentParser"""

    parser = argparse.ArgumentParser(
        description="A script to email students study pages for a semi-open book exam",
        epilog="Note that column count arguments start from zero."

    )

    parser.add_argument('-b', '--batch-mode',
                        action='store_true',
                        dest='batch_mode',
                        default=False,
                        help='run automatically with values given')

    parser.add_argument('--interactive-mode',
                        action='store_true',
                        dest='interactive_mode',
                        default=True,
                        help='prompt the user for details (default)')

    parser.add_argument('-i', '--input-file',
                        dest='input_file',
                        default='students.csv',
                        help='the name of the input CSV file with one row per student')

    parser.add_argument('-sidc', '--student-id-column',
                        dest='student_id_column',
                        default=1,
                        help='the column containing the student id (default 1)')

    parser.add_argument('-snc', '--student-name-column',
                        dest='student_name_column',
                        default=2,
                        help='the column containing the student name (default 2)')

    parser.add_argument('-sec', '--student-email-column',
                        dest='student_email_column',
                        default=9,
                        help='the column containing the student email (default 9)')

    parser.add_argument('-sidregexp', '--student-id-regexp',
                        dest='student_id_regexp',
                        default='B[0-9]+',
                        help='a regular expression for valid student IDs (default B[0-9]+)')

    parser.add_argument('--smtp-server',
                        dest='smtp_server',
                        default='localhost',
                        help='the address of an smtp server')

    parser.add_argument('--email-subject',
                        dest='email_subject',
                        default='IMPORTANT: Your semi-open-book Guide Sheet',
                        help='the subject of emails that are sent')

    parser.add_argument('--email-sender',
                        dest='email_sender',
                        default='noreply@nowhere.org',
                        help='the sender address from which to send emails')

    parser.add_argument('-t', '--test-only',
                        action='store_true',
                        dest='test_only',
                        default=False,
                        help='do not send any emails')

    args = parser.parse_args()

    # Allow for any overrides from program logic or interaction with the user
    args = override_arguments(args)
    return(args)


def main():
    """the main function that kicks everything else off"""

    print("Hello")
    args = parse_arguments()

    print("Starting open-book...")
    print(args)
    csvReader = csv.reader(open(args.input_file, 'r'), dialect='excel')

    student_count = 0
    # Go through each row
    for row in csvReader:
        student_number = row[args.student_id_column]
        # Check if the second cell looks like a student number
        if re.match(args.student_id_regexp, row[args.student_id_column]):
            student_count = student_count + 1
            process_student(args, row)
        else:
            print('  Skipping: non matching row')

    print('Stopping open-book...')


if __name__ == '__main__':
    main()
