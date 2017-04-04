# node-read-write-sheet
Node.js sample application that loads a sheet, updates selected cells, and saves the results

This is a minimal Smartsheet sample that demonstrates how to
* Load a sheet
* Loop through the rows
* Check for rows that meet a criteria
* Update cell values
* Write the results back to the original sheet


This sample scans a sheet for rows where the value of the "Status" column is "Complete" and sets the "Remaining" column to zero.
This is implemented in the `evaluate_row_and_build_updates()` method which you should modify to meet your needs.


## Setup
* Install the smartsheet library with `npm install smartsheet` at the command line
* Import the sample data from "Sample Sheet.xlsx" into a new sheet

* Update the `config.json` file with these two settings:
    * An API access token, obtained from the Smartsheet Account button, under Personal settings
    * The Sheet Id, obtained from sheet properties 

* Run the application using your preferred IDE or at the command line with `node node-read-write-sheet.js` 

The rows marked "Complete" will have the "Remaining" value set to 0. (Note that you will have to refresh in the desktop application to see the changes)

## See also
- http://smartsheet-platform.github.io/api-docs/
- https://github.com/smartsheet-platform/smartsheet-javascript-sdk
- https://www.smartsheet.com/
