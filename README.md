# VBA Concatenation and Deletion
Excel script to concatenate data into a semicolon delimited column, then delete the duplicate rows.

### Use Case
My use case for this script is to dedupe a list of companies and their email domains, by concatenating all the email domains into one field so the duplicate rows can be deleted.

## How to Run
* Open a file in Excel (test file of dummy data included in this repo) and go to your developer tools to add a VBA module.
* Copy `code.vba` and note that setting the new column header in line 3 may belong in a different column. My dummy data only has 3 columns, so 4 is sufficient.
* Select all the data in your Excel file before you run the script. We are calculating the rows for the variable `RCount` based on your selection, so this is important.
* Run the script!

Using the dummy data example, here is what the file looks like initially:

![Initial Data](/StartingFile.JPG "Initial Data")

After the script runs, we end with:

![Final Data](/EndingFile.JPG "Final Data")