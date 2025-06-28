function onEdit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  
  // Get the actual data range (from row 25 where headers are)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(25, 1, lastRow - 24, lastCol).getValues();
  
  const headers = data[0];
  const itemCol = headers.indexOf("Item");
  const qtyCol = headers.indexOf("Quantity (estimated)");
  const dailyUseCol = headers.indexOf("How much we use per day (estimate)");
  const stationCol = headers.indexOf("Item in each station?");
  const unitCol = headers.indexOf("Unit");

  // Debug: Log the headers and column indices
  console.log("Headers found:", headers);
  console.log("Item column index:", itemCol);
  console.log("Quantity column index:", qtyCol);
  console.log("Daily use column index:", dailyUseCol);
  console.log("Station column index:", stationCol);
  console.log("Unit column index:", unitCol);

  // Check if required columns exist
  if (itemCol === -1 || qtyCol === -1 || dailyUseCol === -1) {
    console.error("Required columns not found. Please check column names.");
    return;
  }

  // Add "Days of Supply Left" header to column F (column 6)
  const daysSupplyHeader = "Days of Supply Left";
  sheet.getRange(25, 6).setValue(daysSupplyHeader);
  
  // Style the header
  sheet.getRange(25, 6).setFontWeight("bold");
  sheet.getRange(25, 6).setBackground("#E8EAF6");

  const recipients = ["email@email.com", "email@email.com"]; 

  let lowItems = [];
  let redLevelItems = []; // Items with less than 10 days
  let allItems = []; // All items for Sheet2

  for (let i = 1; i < data.length; i++) {
    const item = data[i][itemCol];
    const qty = parseFloat(data[i][qtyCol]);
    const dailyUse = parseFloat(data[i][dailyUseCol]);
    const stationItem = data[i][stationCol];

    // Skip rows with missing data
    if (!item || isNaN(qty) || isNaN(dailyUse) || dailyUse <= 0) {
      continue;
    }

    let daysLeft = qty / dailyUse;
    
    // Debug: Log the calculation for Green Soap specifically
    if (item === "Green Soap") {
      console.log("Green Soap calculation:");
      console.log("Quantity:", qty);
      console.log("Daily use:", dailyUse);
      console.log("Base days left:", daysLeft);
      console.log("Station item value:", stationItem);
      console.log("Station item type:", typeof stationItem);
    }
    
    // Add 5 days if item is in each station (has "1" in station column)
    if (stationItem === 1 || stationItem === "1" || stationItem === 1.0) {
      daysLeft += 5;
      // Debug: Log when 5 days are added
      if (item === "Green Soap") {
        console.log("Added 5 days for station item. New total:", daysLeft);
      }
    }

    const daysLeftFormatted = Math.round(daysLeft * 10) / 10; // Round to 1 decimal place

    // Set the days of supply left value in column F
    sheet.getRange(i + 25, 6).setValue(daysLeftFormatted);

    // Conditional formatting color logic for Quantity column (original column):
    const qtyCell = sheet.getRange(i + 25, qtyCol + 1);
    const unitCell = sheet.getRange(i + 25, unitCol + 1);
    
    if (daysLeft < 10) {
      qtyCell.setBackground("#FF9999"); // Red
      unitCell.setBackground("#FF9999"); // Red - match quantity column
    } else if (daysLeft < 30) {
      qtyCell.setBackground("#FFF59D"); // Yellow
      unitCell.setBackground("#FFF59D"); // Yellow - match quantity column
    } else {
      qtyCell.setBackground("#C8E6C9"); // Green
      unitCell.setBackground("#C8E6C9"); // Green - match quantity column
    }

    // Conditional formatting color logic for Days of Supply Left column (column F):
    const daysCell = sheet.getRange(i + 25, 6);
    if (daysLeft < 10) {
      daysCell.setBackground("#FF9999"); // Red
      daysCell.setFontColor("#FFFFFF"); // White text
    } else if (daysLeft < 30) {
      daysCell.setBackground("#FFF59D"); // Yellow
      daysCell.setFontColor("#000000"); // Black text
    } else {
      daysCell.setBackground("#C8E6C9"); // Green
      daysCell.setFontColor("#000000"); // Black text
    }
    
    // Debug: Log the styling for column F
    if (item === "Green Soap") {
      console.log("Setting column F styling for Green Soap:");
      console.log("Row:", i + 25, "Column: 6");
      console.log("Days left:", daysLeft);
      console.log("Background color:", daysLeft < 10 ? "#FF9999" : daysLeft < 30 ? "#FFF59D" : "#C8E6C9");
    }
    
    if (daysLeft < 10) {
      redLevelItems.push({ item: item, qty: qty, unit: data[i][unitCol], daysLeft: daysLeftFormatted }); // Add to red level items
    }

    // Add to all items for Sheet2
    allItems.push({
      category: data[i][0],
      unit: data[i][unitCol],
      item: item,
      qty: qty,
      dailyUse: dailyUse,
      daysLeft: daysLeftFormatted,
      stationItem: stationItem
    });

    if (daysLeft < 10) {
      lowItems.push(`- ${item}: Only ${qty} left, used ${dailyUse}/day (${daysLeftFormatted} days of stock)`);
    }
  }

  // Update the "We might be running low on:" section (starting from row 3, after headers in row 2)
  updateLowItemsSection(sheet, redLevelItems);

  // Create/update Sheet2 with all items sorted by days of supply left
  createSortedSheet(allItems);

  if (lowItems.length > 0) {
    const message = `🚨 Inventory Alert: The following items are low:\n\n${lowItems.join('\n')}`;
    const subject = "⚠️ Studio Low Inventory Alert";

    recipients.forEach(email => {
      try {
        MailApp.sendEmail(email, subject, message);
      } catch (error) {
        console.error(`Failed to send email to ${email}: ${error.message}`);
      }
    });
  }
}

// Function to update the low items section
function updateLowItemsSection(sheet, redLevelItems) {
  // First, let's find where the low items section actually ends
  // Look for the last row with content in columns C or D, starting from row 3
  let lastLowItemRow = 2; // Start at row 2 (headers)
  
  // Check rows 3-50 to find where the low items section ends
  for (let row = 3; row <= 50; row++) {
    const cellC = sheet.getRange(row, 3).getValue();
    const cellD = sheet.getRange(row, 4).getValue();
    
    if (cellC === "" && cellD === "") {
      // Found empty row, this is where the section ends
      lastLowItemRow = row - 1;
      break;
    }
  }
  
  // If we found content, clear only those specific rows
  if (lastLowItemRow > 2) {
    sheet.getRange(3, 3, lastLowItemRow - 2, 3).clearContent(); // Clear 3 columns now (C, D, E)
  }
  
  // If no red level items, we're done
  if (redLevelItems.length === 0) {
    return;
  }
  
  // Calculate how many rows we need
  const neededRows = redLevelItems.length;
  const startRow = 3; // Start after headers in row 2
  
  // Add rows if necessary (insert after the last row we might need)
  const maxRow = startRow + neededRows - 1;
  const currentLastRow = sheet.getLastRow();
  if (maxRow > currentLastRow) {
    const rowsToAdd = maxRow - currentLastRow;
    sheet.insertRowsAfter(currentLastRow, rowsToAdd);
  }
  
  // Add "Days of Supply Left" header to column E in the low items section
  sheet.getRange(2, 5).setValue("Days of Supply Left");
  sheet.getRange(2, 5).setFontWeight("bold");
  sheet.getRange(2, 5).setBackground("#E8EAF6");
  
  // Populate the items starting from row 3
  for (let i = 0; i < redLevelItems.length; i++) {
    const row = startRow + i;
    const item = redLevelItems[i];
    
    // Set item name in column C
    sheet.getRange(row, 3).setValue(item.item);
    
    // Set quantity and unit in column D
    sheet.getRange(row, 4).setValue(`${item.qty} ${item.unit}`);
    
    // Set days of supply left in column E
    sheet.getRange(row, 5).setValue(item.daysLeft);
    
    // Style the cells
    sheet.getRange(row, 3, 1, 3).setBackground("#FF9999"); // Red background for all 3 columns
    sheet.getRange(row, 3, 1, 3).setFontColor("#FFFFFF"); // White text
    sheet.getRange(row, 3, 1, 3).setFontWeight("bold");
  }
}

// Function to create/update Sheet2 with all items sorted by days of supply left
function createSortedSheet(allItems) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet2 = spreadsheet.getSheetByName("Sheet2");
  
  // Create Sheet2 if it doesn't exist
  if (!sheet2) {
    sheet2 = spreadsheet.insertSheet("Sheet2");
  }
  
  // Clear existing content
  sheet2.clear();
  
  // Sort items by days of supply left (lowest first)
  allItems.sort((a, b) => a.daysLeft - b.daysLeft);
  
  // Create headers (swapped B and D: Unit and Quantity)
  const headers = ["Category", "Quantity (estimated)", "Item", "Unit", "Daily Usage", "Days of Supply Left", "Item in each station?"];
  sheet2.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Style headers
  sheet2.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet2.getRange(1, 1, 1, headers.length).setBackground("#E8EAF6");
  
  // Populate data (swapped B and D: Unit and Quantity)
  const data = allItems.map(item => [
    item.category,
    item.qty,
    item.item,
    item.unit,
    item.dailyUse,
    item.daysLeft,
    item.stationItem
  ]);
  
  // Debug: Log the first few rows to verify data structure
  console.log("Sheet2 data structure:");
  console.log("Headers:", headers);
  if (data.length > 0) {
    console.log("First row data:", data[0]);
    console.log("Days of Supply Left value:", data[0][5]);
  }
  
  if (data.length > 0) {
    sheet2.getRange(2, 1, data.length, headers.length).setValues(data);
    
    // Apply color coding based on days of supply left
    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const daysLeft = data[i][5]; // Days of Supply Left column (still at index 5)
      
      // Debug: Log color coding for first few items
      if (i < 3) {
        console.log(`Row ${i + 1}: ${data[i][2]} (${data[i][1]} ${data[i][3]}) - Days: ${daysLeft}`);
      }
      
      if (daysLeft < 10) {
        sheet2.getRange(row, 1, 1, headers.length).setBackground("#FF9999"); // Red
        sheet2.getRange(row, 1, 1, headers.length).setFontColor("#FFFFFF"); // White text
      } else if (daysLeft < 30) {
        sheet2.getRange(row, 1, 1, headers.length).setBackground("#FFF59D"); // Yellow
        sheet2.getRange(row, 1, 1, headers.length).setFontColor("#000000"); // Black text
      } else {
        sheet2.getRange(row, 1, 1, headers.length).setBackground("#C8E6C9"); // Green
        sheet2.getRange(row, 1, 1, headers.length).setFontColor("#000000"); // Black text
      }
    }
  }
  
  // Auto-resize columns
  sheet2.autoResizeColumns(1, headers.length);
}

// Manual trigger function for testing
function testInventoryCheck() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  
  // Get the actual data range (from row 25 where headers are)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(25, 1, lastRow - 24, lastCol).getValues();
  
  const headers = data[0];
  const itemCol = headers.indexOf("Item");
  const qtyCol = headers.indexOf("Quantity (estimated)");
  const dailyUseCol = headers.indexOf("How much we use per day (estimate)");
  const stationCol = headers.indexOf("Item in each station?");
  const unitCol = headers.indexOf("Unit");

  // Debug: Log the headers and column indices
  console.log("Headers found:", headers);
  console.log("Item column index:", itemCol);
  console.log("Quantity column index:", qtyCol);
  console.log("Daily use column index:", dailyUseCol);
  console.log("Station column index:", stationCol);
  console.log("Unit column index:", unitCol);

  // Check if required columns exist
  if (itemCol === -1 || qtyCol === -1 || dailyUseCol === -1) {
    console.error("Required columns not found. Please check column names.");
    return;
  }

  // Add "Days of Supply Left" header to column F (column 6)
  const daysSupplyHeader = "Days of Supply Left";
  sheet.getRange(25, 6).setValue(daysSupplyHeader);
  
  // Style the header
  sheet.getRange(25, 6).setFontWeight("bold");
  sheet.getRange(25, 6).setBackground("#E8EAF6");

  const recipients = ["email@example.com", "email@example.com"]; 

  let lowItems = [];
  let redLevelItems = []; // Items with less than 10 days
  let allItems = []; // All items for Sheet2

  for (let i = 1; i < data.length; i++) {
    const item = data[i][itemCol];
    const qty = parseFloat(data[i][qtyCol]);
    const dailyUse = parseFloat(data[i][dailyUseCol]);
    const stationItem = data[i][stationCol];

    // Skip rows with missing data
    if (!item || isNaN(qty) || isNaN(dailyUse) || dailyUse <= 0) {
      continue;
    }

    let daysLeft = qty / dailyUse;
    
    // Debug: Log the calculation for Green Soap specifically
    if (item === "Green Soap") {
      console.log("Green Soap calculation:");
      console.log("Quantity:", qty);
      console.log("Daily use:", dailyUse);
      console.log("Base days left:", daysLeft);
      console.log("Station item value:", stationItem);
      console.log("Station item type:", typeof stationItem);
    }
    
    // Add 5 days if item is in each station (has "1" in station column)
    if (stationItem === 1 || stationItem === "1" || stationItem === 1.0) {
      daysLeft += 5;
      // Debug: Log when 5 days are added
      if (item === "Green Soap") {
        console.log("Added 5 days for station item. New total:", daysLeft);
      }
    }

    const daysLeftFormatted = Math.round(daysLeft * 10) / 10; // Round to 1 decimal place

    // Set the days of supply left value in column F
    sheet.getRange(i + 25, 6).setValue(daysLeftFormatted);

    // Conditional formatting color logic for Quantity column (original column):
    const qtyCell = sheet.getRange(i + 25, qtyCol + 1);
    const unitCell = sheet.getRange(i + 25, unitCol + 1);
    
    if (daysLeft < 10) {
      qtyCell.setBackground("#FF9999"); // Red
      unitCell.setBackground("#FF9999"); // Red - match quantity column
    } else if (daysLeft < 30) {
      qtyCell.setBackground("#FFF59D"); // Yellow
      unitCell.setBackground("#FFF59D"); // Yellow - match quantity column
    } else {
      qtyCell.setBackground("#C8E6C9"); // Green
      unitCell.setBackground("#C8E6C9"); // Green - match quantity column
    }

    // Conditional formatting color logic for Days of Supply Left column (column F):
    const daysCell = sheet.getRange(i + 25, 6);
    if (daysLeft < 10) {
      daysCell.setBackground("#FF9999"); // Red
      daysCell.setFontColor("#FFFFFF"); // White text
    } else if (daysLeft < 30) {
      daysCell.setBackground("#FFF59D"); // Yellow
      daysCell.setFontColor("#000000"); // Black text
    } else {
      daysCell.setBackground("#C8E6C9"); // Green
      daysCell.setFontColor("#000000"); // Black text
    }
    
    // Debug: Log the styling for column F
    if (item === "Green Soap") {
      console.log("Setting column F styling for Green Soap:");
      console.log("Row:", i + 25, "Column: 6");
      console.log("Days left:", daysLeft);
      console.log("Background color:", daysLeft < 10 ? "#FF9999" : daysLeft < 30 ? "#FFF59D" : "#C8E6C9");
    }
    
    if (daysLeft < 10) {
      redLevelItems.push({ item: item, qty: qty, unit: data[i][unitCol], daysLeft: daysLeftFormatted }); // Add to red level items
    }

    // Add to all items for Sheet2
    allItems.push({
      category: data[i][0],
      unit: data[i][unitCol],
      item: item,
      qty: qty,
      dailyUse: dailyUse,
      daysLeft: daysLeftFormatted,
      stationItem: stationItem
    });

    if (daysLeft < 10) {
      lowItems.push(`- ${item}: Only ${qty} left, used ${dailyUse}/day (${daysLeftFormatted} days of stock)`);
    }
  }

  // Update the "We might be running low on:" section (starting from row 3, after headers in row 2)
  updateLowItemsSection(sheet, redLevelItems);

  // Create/update Sheet2 with all items sorted by days of supply left
  createSortedSheet(allItems);

  if (lowItems.length > 0) {
    const message = `🚨 Inventory Alert: The following items are low:\n\n${lowItems.join('\n')}`;
    const subject = "⚠️ Studio Low Inventory Alert";

    recipients.forEach(email => {
      try {
        MailApp.sendEmail(email, subject, message);
      } catch (error) {
        console.error(`Failed to send email to ${email}: ${error.message}`);
      }
    });
  }
}

// Function to decrement quantities by daily usage (run daily)
function dailyInventoryUpdate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  
  // Get the actual data range (from row 25 where headers are)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(25, 1, lastRow - 24, lastCol).getValues();
  
  const headers = data[0];
  const itemCol = headers.indexOf("Item");
  const qtyCol = headers.indexOf("Quantity (estimated)");
  const dailyUseCol = headers.indexOf("How much we use per day (estimate)");
  const stationCol = headers.indexOf("Item in each station?");
  const unitCol = headers.indexOf("Unit");

  // Check if required columns exist
  if (itemCol === -1 || qtyCol === -1 || dailyUseCol === -1) {
    console.error("Required columns not found. Please check column names.");
    return;
  }

  let updatedItems = [];
  let zeroItems = [];

  for (let i = 1; i < data.length; i++) {
    const item = data[i][itemCol];
    const qty = parseFloat(data[i][qtyCol]);
    const dailyUse = parseFloat(data[i][dailyUseCol]);

    // Skip rows with missing data
    if (!item || isNaN(qty) || isNaN(dailyUse) || dailyUse <= 0) {
      continue;
    }

    // Calculate new quantity after daily usage
    const newQty = Math.max(0, qty - dailyUse); // Don't go below 0
    
    // Update the quantity in the spreadsheet
    sheet.getRange(i + 25, qtyCol + 1).setValue(newQty);
    
    // Track changes
    updatedItems.push({
      item: item,
      oldQty: qty,
      newQty: newQty,
      dailyUse: dailyUse,
      unit: data[i][unitCol] || "units"
    });
    
    // Track items that reached zero
    if (newQty === 0 && qty > 0) {
      zeroItems.push(item);
    }
  }

  // Log the daily update
  console.log(`Daily inventory update completed at ${new Date().toLocaleString()}`);
  console.log(`Updated ${updatedItems.length} items`);
  
  // Log items that reached zero
  if (zeroItems.length > 0) {
    console.log(`Items that reached zero: ${zeroItems.join(', ')}`);
  }

  // Send email notification for items that reached zero
  if (zeroItems.length > 0) {
    const recipients = ["example@example.com", "example@example.com"];
    const message = `🚨 Daily Inventory Update: The following items have reached zero:\n\n${zeroItems.map(item => `- ${item}`).join('\n')}\n\nPlease restock these items immediately.`;
    const subject = "⚠️ Studio Daily Update - Items Out of Stock";

    recipients.forEach(email => {
      try {
        MailApp.sendEmail(email, subject, message);
      } catch (error) {
        console.error(`Failed to send email to ${email}: ${error.message}`);
      }
    });
  }

  return {
    updatedCount: updatedItems.length,
    zeroItems: zeroItems,
    details: updatedItems
  };
}

// Weekly inventory check function (can be triggered weekly)
function weeklyInventoryCheck() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  
  // Get the actual data range (from row 25 where headers are)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(25, 1, lastRow - 24, lastCol).getValues();
  
  const headers = data[0];
  const itemCol = headers.indexOf("Item");
  const qtyCol = headers.indexOf("Quantity (estimated)");
  const dailyUseCol = headers.indexOf("How much we use per day (estimate)");
  const stationCol = headers.indexOf("Item in each station?");
  const unitCol = headers.indexOf("Unit");

  // Check if required columns exist
  if (itemCol === -1 || qtyCol === -1 || dailyUseCol === -1) {
    console.error("Required columns not found. Please check column names.");
    return;
  }

  const recipients = ["email@email.com", "email@email.com"]; 

  let lowItems = [];
  let redLevelItems = []; // Items with less than 10 days

  for (let i = 1; i < data.length; i++) {
    const item = data[i][itemCol];
    const qty = parseFloat(data[i][qtyCol]);
    const dailyUse = parseFloat(data[i][dailyUseCol]);
    const stationItem = data[i][stationCol];

    // Skip rows with missing data
    if (!item || isNaN(qty) || isNaN(dailyUse) || dailyUse <= 0) {
      continue;
    }

    let daysLeft = qty / dailyUse;
    
    // Add 5 days if item is in each station (has "1" in station column)
    if (stationItem === 1 || stationItem === "1" || stationItem === 1.0) {
      daysLeft += 5;
    }

    const daysLeftFormatted = Math.round(daysLeft * 10) / 10; // Round to 1 decimal place

    if (daysLeft < 10) {
      redLevelItems.push({ item: item, qty: qty, unit: data[i][unitCol], daysLeft: daysLeftFormatted });
      lowItems.push(`- ${item}: Only ${qty} ${data[i][unitCol]} left, used ${dailyUse}/day (${daysLeftFormatted} days of stock)`);
    }
  }

  // Always send weekly report, even if no low items
  const currentDate = new Date().toLocaleDateString();
  let message = `📊 Weekly Inventory Report - ${currentDate}\n\n`;
  
  if (lowItems.length > 0) {
    message += `🚨 The following items are low (less than 10 days of stock):\n\n${lowItems.join('\n')}\n\n`;
  } else {
    message += `✅ All items have sufficient stock (10+ days remaining).\n\n`;
  }
  
  message += `Total items checked: ${data.length - 1}\n`;
  message += `Items requiring attention: ${lowItems.length}`;

  const subject = lowItems.length > 0 ? 
    "⚠️ Studio Weekly Inventory Alert - Low Stock Items" : 
    "✅ Studio Weekly Inventory Report - All Good";

  recipients.forEach(email => {
    try {
      MailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send email to ${email}: ${error.message}`);
    }
  });

  console.log(`Weekly inventory check completed at ${new Date().toLocaleString()}`);
  console.log(`Items checked: ${data.length - 1}, Low items: ${lowItems.length}`);
  
  return {
    itemsChecked: data.length - 1,
    lowItems: lowItems.length,
    details: lowItems
  };
}
