function onOpen() {
	// Add menu button when sheet is opened.
	SpreadsheetApp.getUi()
		.createMenu("345 Sheridan Scripts")
		.addItem("Generate Balances", "main")
		.addToUi();
}

function sortSpreadsheet(expensesSheet, paymentsSheet) {
	expensesSheet.sort(1);
	paymentsSheet.sort(1);
}

function getDisplayMessage(rOwesM) {
	// Round to 2 decimal places.
	const rounded = Math.round(rOwesM * 100) / 100;

	let message = "Current Status: ";

	if (rounded > 0) {
		message += `Rohan owes Marcus $${rounded}.`;
	} else if (rounded < 0) {
		message += `Marcus owes Rohan $${-rounded}.`;
	} else {
		message += "No money is owed.";
	}

	return message;
}

function getExpensesBalance(expensesData) {
	let rOwesM = 0;

	// Skip first row because it's a header row.
	for (const expenseRow of expensesData.slice(1)) {
		const purchaser = expenseRow[2];
		const cost = expenseRow[3];
		const costIsSplit = expenseRow[4];
		const amountToPay = costIsSplit ? cost / 2 : cost;

		if (purchaser === "Rohan") {
			rOwesM -= amountToPay;
		} else if (purchaser === "Marcus") {
			rOwesM += amountToPay;
		} else if (purchaser === "" && cost === "") {
			// Reached empty row. Required for this sheet due to checkboxes.
			break;
		} else {
			throw Error(`Incorrect purchaser in Expenses page: ${purchaser}.`);
		}
	}

	return rOwesM;
}

function getPaymentsBalance(paymentsData) {
	let rOwesM = 0;

	// Skip first row because it's a header row.
	for (const paymentRow of paymentsData.slice(1)) {
		const senderName = paymentRow[1];
		const amount = paymentRow[2];

		if (senderName === "Rohan") {
			rOwesM -= amount;
		} else if (senderName === "Marcus") {
			rOwesM += amount;
		} else {
			throw Error(
				`Incorrect sender name in Payments page: ${senderName}.`
			);
		}
	}

	return rOwesM;
}

function main() {
	const spreadsheet = SpreadsheetApp.getActive();
	const expensesSheet = spreadsheet.getSheetByName("Expenses");
	const paymentsSheet = spreadsheet.getSheetByName("Payments");

	const expensesData = expensesSheet.getDataRange().getValues();
	const paymentsData = paymentsSheet.getDataRange().getValues();

	const rOwesM =
		getExpensesBalance(expensesData) + getPaymentsBalance(paymentsData);
	const displayMessage = getDisplayMessage(rOwesM);
	expensesSheet.getRange("G1").setValue(displayMessage);

	sortSpreadsheet(expensesSheet, paymentsSheet);
}
