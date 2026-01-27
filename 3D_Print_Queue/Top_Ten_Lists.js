/** Made by Chadwick Hobgood, 1/26/2025
 * * Finds Top 10 Requestors (Total) and Top 10 Flagged Requestors.
 * * Maps emails to names and outputs to "Dashboard_Data_Link" Column N:Q.
 */
function updateTopTenDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName("Archive");
  const dashboardSheet = ss.getSheetByName("Dashboard_Data_Link");
  const userSheet = ss.getSheetByName("Users"); // Assumes a sheet with Email (A) and Name (B)

  const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 11).getValues();
  
  // 1. Create a Name Lookup Map {email: name}
  const userMap = {};
  if (userSheet) {
    userSheet.getRange(1, 1, userSheet.getLastRow(), 2).getValues().forEach(row => {
      userMap[row[0].toLowerCase()] = row[1];
    });
  }

  // 2. Count Occurrences (Total and Flagged)
  const totalCounts = {};
  const flaggedCounts = {};

  archiveData.forEach(row => {
    const email = row[2] ? row[2].toLowerCase() : null; // Column C
    const status = row[10] ? row[10].toString() : "";   // Column K
    
    if (email) {
      totalCounts[email] = (totalCounts[email] || 0) + 1;
      if (status.includes("Flagged")) {
        flaggedCounts[email] = (flaggedCounts[email] || 0) + 1;
      }
    }
  });

  // 3. Helper function to sort and format Top 10
  const getTopTen = (countObj) => {
    return Object.keys(countObj)
      .map(email => [userMap[email] || email, countObj[email]]) // Use Name if found, else Email
      .sort((a, b) => b[1] - a[1]) // Sort descending
      .slice(0, 10); // Take top 10
  };

  const topTotal = getTopTen(totalCounts);
  const topFlagged = getTopTen(flaggedCounts);

  // 4. Prepare Output Array (Columns N, O, P, Q)
  const finalOutput = [["Top 10 Requestors", "Count", "Top 10 Flagged", "Count"]];
  
  // Build rows up to 10
  for (let i = 0; i < 10; i++) {
    const totalRow = topTotal[i] || ["", ""];
    const flaggedRow = topFlagged[i] || ["", ""];
    finalOutput.push([...totalRow, ...flaggedRow]);
  }

  // 5. Clear and Write to Dashboard_Data_Link starting at Column N (14)
  dashboardSheet.getRange("N:Q").clearContent();
  dashboardSheet.getRange(1, 14, finalOutput.length, 4).setValues(finalOutput);
  
  // Format Header
  dashboardSheet.getRange("N1:Q1").setFontWeight("bold");

  console.log("Top 10 Dashboards updated.");
}
