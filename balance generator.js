// Add menu button.
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("747 Custom Stuff").addItem("Generate Balances", 'generateBalances').addItem("Sort Sheet", 'sortSpreadsheet').addToUi();
}

// Sorts Expenses and Payments sheets.
function sortSpreadsheet() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getSheetByName('Expenses').sort(1);
    spreadsheet.getSheetByName('Payments').sort(1);
}

function generateBalances() {
    // 2d array to hold current balances. JavaScript is annoying.
    // i owes j $(balances[i][j])
    var balances = new Array(4);
    for (var i = 0; i < 4; ++i) {
        balances[i] = [0, 0, 0, 0];
    }

    // Constants to index balances matrix.
    const ROHAN = 0;
    const MICHAEL = 1;
    const TOM = 2;
    const MARCUS = 3;
    var name_to_index = { "Marcus": MARCUS, "Rohan": ROHAN, "Michael": MICHAEL, "Tom": TOM }

    // Access Expenses data.
    var spreadsheet = SpreadsheetApp.getActive();
    var expenses_sheet = spreadsheet.getSheetByName('Expenses');
    var expenses_data = expenses_sheet.getDataRange().getValues();

    // Loop through expenses.
    for (var i = 1; i < expenses_data.length; ++i) {
        var row = expenses_data[i];
        var purchaser_index = name_to_index[row[2]];
        var cost = row[3];

        // Compute cost per person.
        var num_of_purchasers = 0;
        for (var j = 4; j <= 7; ++j) {
            if (row[j] == "*") {
                ++num_of_purchasers;
            }
        }
        var per_person = cost / num_of_purchasers;

        // Update balances per person.
        for (var j = 4; j <= 7; ++j) {
            if (row[j] == "*") {
                balances[j - 4][purchaser_index] += per_person;
            }
        }
    }

    // Access Payments data.
    var payments_sheet = spreadsheet.getSheetByName('Payments');
    var payments_data = payments_sheet.getDataRange().getValues();

    // Loop through payments.
    for (var i = 1; i < payments_data.length; ++i) {
        var row = payments_data[i];
        var sender_index = name_to_index[row[1]];
        var receiver_index = name_to_index[row[2]];
        var amount = row[4];
        balances[receiver_index][sender_index] += amount;
    }

    // Access Balances sheet.
    var balances_sheet = spreadsheet.getSheetByName('Balances');

    // Update Balances Sheet.
    balances_sheet.getRange('C2').setValue(balances[ROHAN][MICHAEL] - balances[MICHAEL][ROHAN]);
    balances_sheet.getRange('D2').setValue(balances[ROHAN][MARCUS] - balances[MARCUS][ROHAN]);
    balances_sheet.getRange('D3').setValue(balances[MICHAEL][MARCUS] - balances[MARCUS][MICHAEL]);
    balances_sheet.getRange('E2').setValue(balances[ROHAN][TOM] - balances[TOM][ROHAN]);
    balances_sheet.getRange('E3').setValue(balances[MICHAEL][TOM] - balances[TOM][MICHAEL]);
    balances_sheet.getRange('E4').setValue(balances[MARCUS][TOM] - balances[TOM][MARCUS]);
}
