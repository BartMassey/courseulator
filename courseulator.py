#!/usr/bin/python3
# Clean up Banweb "new" classlist XLSX export,
# generating student-emails.txt and "new format" students.csv.
# Bart Massey 2021

import csv, datetime, re, sys, warnings

from openpyxl import load_workbook

name_re = re.compile(
    r"([A-Z][-' A-Za-z]*), ([A-Z][-' A-Za-z]*)( ([A-Z])\.)?( \(([^)]*)\))?",
)

section_re = re.compile(
    r"[1-6]\d\d[Pp]?",
)

withdraw_re = re.compile(
    r"[Ww]ithdraw",
)

sheet = sys.argv[1]

with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")
    wb = load_workbook(filename=sheet)
    ws = wb.active
    v = ws.cell(column=1, row=1).value
    if v != "Subject":
        print(f"{section_file}: sheet header error", file=sys.stderr)
        exit(1)

    headers = { c.value : c.column - 1 for (c,) in ws.iter_cols(
        min_row=1,
        max_row=1,
    ) }
    term_header = headers["Term"]
    crn_header = headers["CRN"]
    status_header = headers["Status"]
    title_header = headers["Title"]
    instructor_last_header = headers["Instructor_Last_Name"]
    instructor_first_header = headers["Instructor_First_Name"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        term = row[term_header]
        crn = row[crn_header]
        status = row[status_header]
        if status == "Cancelled":
            continue
        title = row[title_header]
        instructor_last = row[instructor_last_header]
        instructor_first = row[instructor_first_header]
        if not instructor_last or not instructor_first:
            continue
        instructor = instructor_last + ", " + instructor_first
        print(term, crn, status, title, instructor)
