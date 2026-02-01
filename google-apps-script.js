// ═══════════════════════════════════════════════════════════════════
//  ILM Academy - Interview Rubric Google Apps Script
//
//  SETUP INSTRUCTIONS:
//  1. Go to https://sheets.google.com and create a new spreadsheet
//  2. Name it "ILM Academy Interview Scores"
//  3. Go to Extensions > Apps Script
//  4. Delete any existing code and paste this entire file
//  5. Click the disk icon to Save (or Ctrl+S)
//  6. Click "Deploy" > "New deployment"
//  7. Choose type: "Web app"
//  8. Set "Execute as": Me
//  9. Set "Who has access": Anyone
// 10. Click "Deploy" and authorize when prompted
// 11. Copy the Web App URL
// 12. Paste that URL into the SCRIPT_URL variable in index.html
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'Responses';
    var sheet = ss.getSheetByName(sheetName);

    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    var data = JSON.parse(e.postData.contents);

    var headers = [
      'Timestamp',
      'Candidate Name',
      'Interview Date',
      'Committee Member',
      'Role',
      'Q1 Rating', 'Q1 Score',
      'Q2 Rating', 'Q2 Score',
      'Q3 Rating', 'Q3 Score',
      'Q4 Rating', 'Q4 Score',
      'Q5 Rating', 'Q5 Score',
      'Q6 Rating', 'Q6 Score',
      'Q7 Rating', 'Q7 Score',
      'Q8 Rating', 'Q8 Score',
      'Total Score',
      'Q1 Comments', 'Q2 Comments', 'Q3 Comments', 'Q4 Comments',
      'Q5 Comments', 'Q6 Comments', 'Q7 Comments', 'Q8 Comments'
    ];

    // Build a map of header name -> value for this submission
    var dataMap = {
      'Timestamp': data.submittedAt || new Date().toISOString(),
      'Candidate Name': data.candidateName || '',
      'Interview Date': data.interviewDate || '',
      'Committee Member': data.committeeMember || '',
      'Role': data.role || '',
      'Total Score': data.totalScore || 0
    };
    for (var i = 0; i < 8; i++) {
      var q = data.questions[i];
      dataMap['Q' + (i + 1) + ' Rating'] = q.rating || '';
      dataMap['Q' + (i + 1) + ' Score'] = q.score || 0;
      dataMap['Q' + (i + 1) + ' Comments'] = q.comments || '';
    }

    // Ensure header row exists and is correct
    var needsHeaders = false;
    var existingHeaders = [];

    if (sheet.getLastRow() === 0) {
      needsHeaders = true;
    } else {
      var lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }
      // Check if first row looks like a header (contains at least one known header name)
      var knownHeaders = ['Timestamp', 'Candidate Name', 'Total Score'];
      var hasAnyHeader = existingHeaders.some(function(h) {
        return knownHeaders.indexOf(String(h).trim()) !== -1;
      });
      if (!hasAnyHeader) {
        needsHeaders = true;
      }
    }

    if (needsHeaders) {
      // No valid headers found — insert canonical headers
      if (sheet.getLastRow() > 0) {
        sheet.insertRowBefore(1);
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      existingHeaders = headers;

      // Format header row
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#065A5E');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // Check for any missing headers and append them
    var headerIndex = {};
    for (var c = 0; c < existingHeaders.length; c++) {
      headerIndex[String(existingHeaders[c]).trim()] = c;
    }
    var nextCol = existingHeaders.length;
    for (var h = 0; h < headers.length; h++) {
      if (!(headers[h] in headerIndex)) {
        // This expected header is missing — add it at the end
        sheet.getRange(1, nextCol + 1).setValue(headers[h]);
        var cell = sheet.getRange(1, nextCol + 1);
        cell.setFontWeight('bold');
        cell.setBackground('#065A5E');
        cell.setFontColor('#FFFFFF');
        headerIndex[headers[h]] = nextCol;
        nextCol++;
      }
    }

    // Build row array mapped to actual column positions
    var totalCols = nextCol;
    var row = new Array(totalCols);
    for (var c = 0; c < totalCols; c++) {
      row[c] = '';
    }
    for (var headerName in dataMap) {
      if (headerName in headerIndex) {
        row[headerIndex[headerName]] = dataMap[headerName];
      }
    }

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function - run this to verify the script works
function testDoPost() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        candidateName: 'Test Candidate',
        interviewDate: '2026-01-31',
        committeeMember: 'Test Member',
        role: 'Chair',
        questions: [
          { number: 1, maxPoints: 15, rating: 'Exceeds', score: 15, comments: 'Test' },
          { number: 2, maxPoints: 10, rating: 'Meets', score: 8, comments: '' },
          { number: 3, maxPoints: 15, rating: 'Somewhat', score: 7, comments: '' },
          { number: 4, maxPoints: 15, rating: 'Exceeds', score: 15, comments: '' },
          { number: 5, maxPoints: 10, rating: 'Meets', score: 8, comments: '' },
          { number: 6, maxPoints: 10, rating: 'Exceeds', score: 10, comments: '' },
          { number: 7, maxPoints: 10, rating: 'Meets', score: 8, comments: '' },
          { number: 8, maxPoints: 15, rating: 'Exceeds', score: 15, comments: '' }
        ],
        totalScore: 86,
        submittedAt: new Date().toISOString()
      })
    }
  };

  var result = doPost(testData);
  Logger.log(result.getContent());
}
