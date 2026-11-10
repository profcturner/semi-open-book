# semi-open-book

A Python script to produce personalised guide sheets for semi-open-book
exams. Students get a single page, showing their student number (and
optionally a barcode), on which they handwrite notes to bring into the exam.

The script reads a CSV file with one row per student, makes a PDF for each
student, saves the PDFs in a folder, and can email each student their sheet.

## Requirements

* Python 3.8 or later
* A TeX distribution that provides `pdflatex`: [MiKTeX](https://miktex.org/)
  on Windows, or TeX Live on Linux and macOS. The template only uses the
  standard packages `geometry`, `fontenc`, `lmodern` and `makebarcode`;
  MiKTeX installs them automatically the first time if needed.

Earlier versions used `latex`, `dvipdf` and PSTricks (`pst-barcode`). None of
these are needed any more.

## Files

* `open-book.py`: the script to run
* `open_book_common.py`: code shared with other versions of the script; keep it
  in the same folder as `open-book.py`
* `open-book.tex`: the LaTeX template for the sheet. Edit the wording here.
  `\ModuleCode` and `\StudentNumber` are filled in for each student.

## The CSV file

One row per student. Rows whose student ID column does not match the student
ID pattern (by default `B` followed by digits) are skipped, so a header row is
fine. Columns are counted from zero; by default the student ID is in column 1,
the name in column 2 and the email address in column 9. Use `-sidc`, `-snc`
and `-sec` to change these.

## Examples

Run it with no options to be prompted for the main settings:

    python open-book.py

Make PDFs only, without sending any email, for module EEE122:

    python open-book.py -b -t -m EEE122 -i students.csv

The PDFs are saved in the `sheets` folder (change with `-o`), named like
`EEE122-B00123456.pdf`.

Include a barcode under the student number:

    python open-book.py -b -t -m EEE122 --barcode

Make and email the sheets through an SMTP server that needs a login (the
password is prompted for, or read from the `OPEN_BOOK_SMTP_PASSWORD`
environment variable):

    python open-book.py -b -m EEE122 --smtp-server smtp.example.ac.uk \
        --smtp-port 587 --smtp-starttls --smtp-user me@example.ac.uk \
        --email-sender me@example.ac.uk --email-body-file message.txt

Run `python open-book.py --help` for all the options.

If a sheet cannot be made, the script reports the end of the LaTeX log for
that student, carries on with the others, lists the failures at the end and
exits with an error.
