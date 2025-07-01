write a python application that takes a class list from the pdf file and
updates the Excel spreadsheet by adding AIG columns.  It also takes a list of students in a word document and adds those to the excel spreadsheet in the same two reading and math Excel columns. The results should be an Excel
spreadsheet that shows if the student is in AIG math and if the student is in AIG readings.
# Excel
* The first row of each page of the Excel spreadsheet has the grade, track and name of the classroom.
* The second row of each page of the Excel spreadsheet is the heading and can be ignored
* Only use the first two columns of the Excel spreadsheet for input, delete all other columns
# PDF
The PDF contains columns with Name Student Id Grade Reading Math. Use the student name and whether they are in AIG math and/or AIG reading from the PDF.
* The PDF also contains the grade of the student
* The PDF file has names with a comma, the Excel spreadsheet uses two columns for the name, this needs to be merged. The second column is always a number so anything between the comma and the number is the first name
* TD, AG, IG and AIG all are the same as AIG status
# Word
* The word document contains a table with three columns
* The first column in the word document contains a name. If there is a comma in the name then the name is "last, first". Otherwise the name is "first last"
* The other two columns in the Word document contain reading and math AIG. TD means AIG in this table
# Output
* Color the row light blue (ADD8E6) if the student is math
* Color the row orange if the student is reading
* Color the row yellow if the student both math and reading
* Use colors in the excel spreadsheets
* Keep the sheet color in the Excel spreadsheet
* Generate a second Excel spreadsheet that removes all the rows where AIG status is None
* Generate an Excel spreadsheet with a list of students from the PDF and Word document that are not in the Excel spreadsheet
* All excel files should go into a directory called output
# Other stuff
* Use python virtual envionrment is in a directory called venv
* The python packages should be listed in a file called requirements.txt
* Remove all the temporary files leaving only the final Excel spreadsheet
* On standard output list how many students have AIG status
* On standard output list how many students are only TD
