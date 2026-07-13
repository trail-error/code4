# code4
learning_python

function main(workbook: ExcelScript.Workbook, event: ExcelScript.WorkbookChangedArgs) {
  // 1. Verify the change happened on the "Tracker" sheet
  let sheet = workbook.getActiveWorksheet();
  if (sheet.getName() !== "Tracker") {
    return; 
  }

  // 2. Get the range of cells that were just changed
  let changedRange = sheet.getRange(event.address);
  
  // 3. Define your columns (Change these index numbers if needed)
  // Column A = 0, B = 1, C = 2, D = 3, etc.
  // Let's assume "SAC#" is Column C (index 2) and "SAC date" is Column D (index 3)
  const targetColumnIndex = 2; // Column index for "SAC#"
  const dateColumnIndex = 3;   // Column index for "SAC date"

  // Check if the edited cell is in the "SAC#" column
  if (changedRange.getColumnIndex() === targetColumnIndex) {
    let rowCount = changedRange.getRowCount();
    
    // Loop through in case multiple rows were pasted at once
    for (let i = 0; i < rowCount; i++) {
      let cell = changedRange.getCell(i, 0);
      let rowIndex = cell.getRowIndex();
      let sacValue = cell.getValue();
      
      let dateCell = sheet.getCell(rowIndex, dateColumnIndex);
      
      // If the user filled in a SAC#, stamp the date
      if (sacValue !== "") {
        // Only stamp if it's currently empty so we don't overwrite old dates
        if (dateCell.getValue() === "") {
          let today = new Date();
          // Formats as MM/DD/YYYY - adjust locale if you need DD/MM/YYYY
          let dateString = today.toLocaleDateString("en-US"); 
          dateCell.setValue(dateString);
        }
      } else {
        // Optional: If they delete the SAC#, clear the date
        dateCell.setValue("");
      }
    }
  }
}
