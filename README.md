# VBA Concatenation and Deletion
Excel script to concatenate data into a semicolon delimited column, then delete the duplicate rows.

### Use Case
My use case for this script is to dedupe a list of companies and their email domains, by concatenating all the email domains into one field so the duplicate rows can be deleted.

## How to Run
* Open a file in Excel (test file of dummy data included in this repo) and go to your developer tools to add a VBA module.
* Copy `code.vba` and update the settings within the asterisks, setting the new column header in line 6 to your new column. My dummy data only has 3 columns, so 4 is sufficient.
* Select the column of your data you want to concatenate in the new column (line 10)
* Run the script!

Using the dummy data example, here is what the file looks like initially:

![Initial Data](/StartingFile.JPG "Initial Data")

After the script runs, we end with:

![Final Data](/EndingFile.JPG "Final Data")