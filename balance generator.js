// Add menu button.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("424 Landfair Scripts").addItem("Generate Balances", 'generateBalances').addToUi();
}

// Sorts Expenses and Payments sheets.
function sortSpreadsheet() {
    const spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getSheetByName('Expenses').sort(1);
    spreadsheet.getSheetByName('Payments').sort(1);
}

function generateBalances() {
    // 2d array to hold current balances. JavaScript is annoying.
    // i owes j $(balances[i][j])
    const balances = new Array(4);
    for (var i = 0; i < 4; ++i) {
        balances[i] = [0, 0, 0, 0];
    }

    // Constants to index balances matrix.
    const ROHAN = 0;
    const MICHAEL = 1;
    const SARAH = 2;
    const MARCUS = 3;
    const name_to_index = { "Marcus": MARCUS, "Rohan": ROHAN, "Michael": MICHAEL, "Sarah": SARAH }

    // Access Expenses data.
    const spreadsheet = SpreadsheetApp.getActive();
    const expenses_sheet = spreadsheet.getSheetByName('Expenses');
    const expenses_data = expenses_sheet.getDataRange().getValues();

    // Loop through expenses.
    for (var i = 1; i < expenses_data.length; ++i) {
        const row = expenses_data[i];
        const purchaser_index = name_to_index[row[2]];
        const cost = row[3];

        // Compute cost per person.
        let num_of_purchasers = 0;
        for (let j = 4; j <= 7; ++j) {
            if (row[j] == "*") {
                ++num_of_purchasers;
            }
        }
        var per_person = cost / num_of_purchasers;

        // Update balances per person.
        for (let j = 4; j <= 7; ++j) {
            if (row[j] == "*") {
                balances[j - 4][purchaser_index] += per_person;
            }
        }
    }

    // Access Payments data.
    const payments_sheet = spreadsheet.getSheetByName('Payments');
    const payments_data = payments_sheet.getDataRange().getValues();

    // Loop through payments.
    for (let i = 1; i < payments_data.length; ++i) {
        const row = payments_data[i];
        const sender_index = name_to_index[row[1]];
        const receiver_index = name_to_index[row[2]];
        const amount = row[4];
        balances[receiver_index][sender_index] += amount;
    }

    // Access Balances sheet.
    const balances_sheet = spreadsheet.getSheetByName('Balances');

    // Update Balances Sheet.
    balances_sheet.getRange('C2').setValue(balances[ROHAN][MICHAEL] - balances[MICHAEL][ROHAN]);
    balances_sheet.getRange('D2').setValue(balances[ROHAN][MARCUS] - balances[MARCUS][ROHAN]);
    balances_sheet.getRange('D3').setValue(balances[MICHAEL][MARCUS] - balances[MARCUS][MICHAEL]);
    balances_sheet.getRange('E2').setValue(balances[ROHAN][SARAH] - balances[SARAH][ROHAN]);
    balances_sheet.getRange('E3').setValue(balances[MICHAEL][SARAH] - balances[SARAH][MICHAEL]);
    balances_sheet.getRange('E4').setValue(balances[MARCUS][SARAH] - balances[SARAH][MARCUS]);

    const messages = [];

    // const names = { MARCUS, ROHAN, MICHAEL, SARAH }
    const names = [ "Marcus", "Rohan", "Michael", "Sarah" ]

    for (let i = 0; i < names.length-1; ++i) {
        for (let j = i+1; j < names.length; ++j) {
            const name1 = names[i];
            const name2 = names[j];

            const val1 = name_to_index[name1]
            const val2 = name_to_index[name2]

            const p1_owes_p2 = balances[val1][val2] - balances[val2][val1];

            if (p1_owes_p2 != 0){
                let message = "!!!";
                const money = Number(p1_owes_p2).toFixed(2);
                if (p1_owes_p2 > 0) {
                    message = `${name1} owes ${name2} $${money}.`
                } else {
                    message = `${name2} owes ${name1} $${-money}.`
                }

                messages.push(message);
            }
        }
    }

    const new_balances_sheet = spreadsheet.getSheetByName('Balances 2.0');
    for (let i = 0; i < messages.length; ++i) {
        new_balances_sheet.getRange(`A${i+1}`).setValue(messages[i]);
    }

    sortSpreadsheet()
}
