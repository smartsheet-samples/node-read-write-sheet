console.log('Starting');

// If not found, install SDK package with command line: npm install smartsheet
var client = require('smartsheet');

// The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
var columnMap = {};

// Helper function to find cell in a row
function getCellByColumnName(row, columnName) {
    var columnId = columnMap[columnName];
    return row.cells.find(function(c) {
        return (c.columnId == columnId);
    });
}

// TODO: Replace the body of this function with your code
// This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
//
// Return a new Row with updated cell values, else null to leave unchanged
function evaluateRowAndBuildUpdates(sourceRow) {
    var rowToUpdate = null;

    // Find the cell and value to evaluate
    var statusCell = getCellByColumnName(sourceRow, "Status");
    if (statusCell.displayValue == "Complete") {
        var remainingCell = getCellByColumnName(sourceRow, "Remaining");
        if (remainingCell.displayValue != "0") { // Skip if already 0
            console.log("Need to update row # " + sourceRow.rowNumber);

            // Build updated row with new cell value
            rowToUpdate = {
                id: sourceRow.id,
                cells: [{
                    columnId: columnMap["Remaining"],
                    value: 0
                }]
            };
        }
    }
    return rowToUpdate;
}

// Initialize client SDK. Uses API token from environment variable "SMARTSHEET_ACCESS_TOKEN"
var ss = client.createClient({ logLevel: 'info' });

var options = {
    path: "Sample Sheet.xlsx",
    fileName: "Sample Sheet.xlsx",
    queryParameters: {
        sheetName: "Sample Sheet",
        headerRowIndex: 0
    }
};

ss.sheets.importXlsxSheet(options)
    .then(function(result) {
        console.log("Created sheet '" + result.result.id + "' from excel file");
        // Load entire sheet
        ss.sheets.getSheet({ id: result.result.id })
        .then(function(sheet) {
            console.log("Loaded " + sheet.rows.length + " rows from sheet '" + sheet.name + "'");

            // Build column map for later reference - converts column name to column id
            sheet.columns.forEach(function(column) {
                columnMap[column.title] = column.id;
            });

            // Accumulate rows needing update here
            var rowsToUpdate = [];

            // Evaluate each row in sheet
            sheet.rows.forEach(function(row) {
                var rowToUpdate = evaluateRowAndBuildUpdates(row);
                if (rowToUpdate)
                    rowsToUpdate.push(rowToUpdate);
            });

            if (rowsToUpdate.length == 0) {
                console.log("No updates required");
            } else {
                // Finally, write all updated cells back to Smartsheet
                console.log("Writing " + rowsToUpdate.length + " rows back to sheet id " + sheet.id);

                var updateRowArgs = {
                    body: rowsToUpdate,
                    sheetId: sheet.id
                };

                ss.sheets.updateRow(updateRowArgs)
                    .then(function(updatedRows) {
                        console.log("Updated succeded");
                    })
                    .catch(function(error) {
                        console.log(error);
                    });

            }
            console.log("Done");
        })
        .catch(function(error) {
            console.log(error);
        });
    })
    .catch(function(error) {
        console.log(error);
    });
