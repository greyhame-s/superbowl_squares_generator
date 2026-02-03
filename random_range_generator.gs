function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Random Generator')
    .addItem('Generate Random Numbers', 'generateScores')
    .addToUi();
}

function generateScores() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  // Validate exactly 10 cells
  if (range.getNumRows() * range.getNumColumns() !== 10) {
    SpreadsheetApp.getUi().alert('Please select exactly 10 cells');
    return;
  }
  
  // Validate cells are empty
  const values = range.getValues();
  const isEmpty = values.every(row => row.every(cell => cell === ''));
  if (!isEmpty) {
    SpreadsheetApp.getUi().alert('Selected cells must be empty');
    return;
  }
  
  randomNumArray = generateRandom();
  
  // Populate cells in sequence
  //const outputValues = randomNumArray.map(val => [val]);
  
  let outputValues;
  if (range.getNumRows() === 1) {
    // Single row, multiple columns - flatten to 1D array wrapped in array
    outputValues = [randomNumArray];
  } else {
    // Multiple rows, single column - keep as 2D array
    outputValues = randomNumArray.map(val => [val]);
  }
  range.setValues(outputValues);
  // Apply formatting
  range.setBackground('#E8F5E9');
  range.setFontColor('#1B5E20');
  range.setFontWeight('bold');
  range.setBorder(true, true, true, true, false, false, '#4CAF50', SpreadsheetApp.BorderStyle.SOLID);
  
  // Audit cell: 12 columns to the right of the first cell in range
  const auditCell = sheet.getRange(range.getRow(), range.getColumn() + 12);
  auditCell.setValue(`Generated: ${new Date()} by ${Session.getEffectiveUser().getEmail()}`);
  auditCell.setFontSize(9);
  auditCell.setFontColor('#999999');
  
  // Lock the generated range permanently
  const protection = range.protect();
  protection.setDescription('Generated random numbers - permanently locked');
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  // Lock the audit cell permanently
  const auditProtection = auditCell.protect();
  auditProtection.setDescription('Audit log - permanently locked');
  auditProtection.removeEditors(protection.getEditors());
  if (auditProtection.canDomainEdit()) {
    auditProtection.setDomainEdit(false);
  }
}

function generateRandom() {
  randomNumArray = new Array(10);
  randomNumArray.fill(-1);
  max = 10;
  genCount = 0;
  while (genCount < max) {
    next = Math.floor(Math.random() * max);
    // Is this number already generated?
    while (randomNumArray.indexOf(next) != -1) {  
      next = Math.floor(Math.random() * max);
    }
    randomNumArray[genCount] = next;
    console.log("Set " + next + " at position " + genCount);
    genCount++;
  }
  return randomNumArray;
}
