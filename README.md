# courseulator: deal with a course report
Bart Massey 2024

This quick Python script tries to figure out the latest
offering of each of our PSU CS 410 "TOPIC" courses. The
input is an Excel spreadsheet.

The basic approach is to try to group similar-looking course
titles in date-sorted order, and produce reports from there.

See `tiles.py` for a string-similarity measurement that has
been tuned for the use case.

# License

This work is licensed under the "MIT License". Please see the file
`LICENSE.txt` in this distribution for license terms.
