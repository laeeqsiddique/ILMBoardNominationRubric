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

// Section config: maps section name to sheet name and header builder
var SECTION_CONFIG = {
  'interview': {
    sheetName: 'Responses',
    getHeaders: function() {
      var h = ['Timestamp', 'Candidate Name', 'Interview Date', 'Committee Member', 'Role'];
      for (var i = 1; i <= 8; i++) {
        h.push('Q' + i + ' Rating', 'Q' + i + ' Score');
      }
      h.push('Total Score');
      for (var i = 1; i <= 8; i++) {
        h.push('Q' + i + ' Comments');
      }
      return h;
    },
    buildDataMap: function(data) {
      var map = {
        'Timestamp': data.submittedAt || new Date().toISOString(),
        'Candidate Name': data.candidateName || '',
        'Interview Date': data.interviewDate || '',
        'Committee Member': data.committeeMember || '',
        'Role': data.role || '',
        'Total Score': data.totalScore || 0
      };
      for (var i = 0; i < 8; i++) {
        var q = data.questions[i];
        map['Q' + (i + 1) + ' Rating'] = q.rating || '';
        map['Q' + (i + 1) + ' Score'] = q.score || 0;
        map['Q' + (i + 1) + ' Comments'] = q.comments || '';
      }
      return map;
    }
  },
  'scenario': {
    sheetName: 'Scenario Responses',
    getHeaders: function() {
      return [
        'Timestamp', 'Candidate Name', 'Interview Date', 'Committee Member', 'Role',
        'Scenario Description',
        'Q1 Rating', 'Q1 Score',
        'Total Score',
        'Q1 Comments'
      ];
    },
    buildDataMap: function(data) {
      var q = data.questions[0];
      return {
        'Timestamp': data.submittedAt || new Date().toISOString(),
        'Candidate Name': data.candidateName || '',
        'Interview Date': data.interviewDate || '',
        'Committee Member': data.committeeMember || '',
        'Role': data.role || '',
        'Scenario Description': data.scenarioDescription || '',
        'Q1 Rating': q.rating || '',
        'Q1 Score': q.score || 0,
        'Total Score': data.totalScore || 0,
        'Q1 Comments': q.comments || ''
      };
    }
  },
  'preinterview': {
    sheetName: 'Pre-Interview Responses',
    getHeaders: function() {
      var h = ['Timestamp', 'Candidate Name', 'Interview Date', 'Committee Member', 'Role'];
      for (var i = 1; i <= 6; i++) {
        h.push('Q' + i + ' Rating', 'Q' + i + ' Score');
      }
      h.push('Total Score');
      for (var i = 1; i <= 6; i++) {
        h.push('Q' + i + ' Comments');
      }
      return h;
    },
    buildDataMap: function(data) {
      var map = {
        'Timestamp': data.submittedAt || new Date().toISOString(),
        'Candidate Name': data.candidateName || '',
        'Interview Date': data.interviewDate || '',
        'Committee Member': data.committeeMember || '',
        'Role': data.role || '',
        'Total Score': data.totalScore || 0
      };
      for (var i = 0; i < 6; i++) {
        var q = data.questions[i];
        map['Q' + (i + 1) + ' Rating'] = q.rating || '';
        map['Q' + (i + 1) + ' Score'] = q.score || 0;
        map['Q' + (i + 1) + ' Comments'] = q.comments || '';
      }
      return map;
    }
  },
  'followup': {
    sheetName: 'Follow-Up Responses',
    getHeaders: function() {
      var h = ['Timestamp', 'Candidate Name', 'Interview Date', 'Committee Member', 'Role'];
      for (var i = 1; i <= 10; i++) {
        h.push('Q' + i + ' Text', 'Q' + i + ' Rating', 'Q' + i + ' Score');
      }
      h.push('Total Score');
      for (var i = 1; i <= 10; i++) {
        h.push('Q' + i + ' Comments');
      }
      return h;
    },
    buildDataMap: function(data) {
      var map = {
        'Timestamp': data.submittedAt || new Date().toISOString(),
        'Candidate Name': data.candidateName || '',
        'Interview Date': data.interviewDate || '',
        'Committee Member': data.committeeMember || '',
        'Role': data.role || '',
        'Total Score': data.totalScore || 0
      };
      for (var i = 0; i < 10; i++) {
        var q = data.questions[i];
        map['Q' + (i + 1) + ' Text'] = q ? (q.questionText || '') : '';
        map['Q' + (i + 1) + ' Rating'] = q ? (q.rating || '') : '';
        map['Q' + (i + 1) + ' Score'] = q ? (q.score || 0) : 0;
        map['Q' + (i + 1) + ' Comments'] = q ? (q.comments || '') : '';
      }
      return map;
    }
  }
};

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // Determine which section (default to 'interview' for backward compatibility)
    var section = data.section || 'interview';
    var config = SECTION_CONFIG[section];
    if (!config) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown section: ' + section }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheetName = config.sheetName;
    var sheet = ss.getSheetByName(sheetName);

    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    var headers = config.getHeaders();
    var dataMap = config.buildDataMap(data);

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
      // Check if first row looks like a header
      var knownHeaders = ['Timestamp', 'Candidate Name', 'Total Score'];
      var hasAnyHeader = existingHeaders.some(function(h) {
        return knownHeaders.indexOf(String(h).trim()) !== -1;
      });
      if (!hasAnyHeader) {
        needsHeaders = true;
      }
    }

    if (needsHeaders) {
      if (sheet.getLastRow() > 0) {
        sheet.insertRowBefore(1);
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      existingHeaders = headers;

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
        section: 'interview',
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

function testScenario() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        section: 'scenario',
        candidateName: 'Test Candidate',
        interviewDate: '2026-01-31',
        committeeMember: 'Test Member',
        role: 'Chair',
        scenarioDescription: 'A parent complains about the curriculum...',
        questions: [
          { number: 1, maxPoints: 10, rating: 'Meets', score: 8, comments: 'Good response' }
        ],
        totalScore: 8,
        submittedAt: new Date().toISOString()
      })
    }
  };

  var result = doPost(testData);
  Logger.log(result.getContent());
}

function testPreInterview() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        section: 'preinterview',
        candidateName: 'Test Candidate',
        interviewDate: '2026-01-31',
        committeeMember: 'Test Member',
        role: '',
        questions: [
          { number: 1, maxPoints: 10, rating: 'Exceeds', score: 10, comments: '' },
          { number: 2, maxPoints: 10, rating: 'Meets', score: 8, comments: '' },
          { number: 3, maxPoints: 10, rating: 'Meets', score: 7, comments: '' },
          { number: 4, maxPoints: 10, rating: 'Somewhat', score: 5, comments: '' },
          { number: 5, maxPoints: 10, rating: 'Exceeds', score: 9, comments: '' },
          { number: 6, maxPoints: 10, rating: 'Meets', score: 8, comments: '' }
        ],
        totalScore: 47,
        submittedAt: new Date().toISOString()
      })
    }
  };

  var result = doPost(testData);
  Logger.log(result.getContent());
}

function testFollowUp() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        section: 'followup',
        candidateName: 'Test Candidate',
        interviewDate: '2026-01-31',
        committeeMember: 'Test Member',
        role: '',
        questions: [
          { number: 1, maxPoints: 10, questionText: 'How do you handle conflict?', rating: 'Exceeds', score: 10, comments: '' },
          { number: 2, maxPoints: 10, questionText: 'What is your time commitment?', rating: 'Meets', score: 8, comments: '' },
          { number: 3, maxPoints: 10, questionText: '', rating: 'Somewhat', score: 5, comments: '' },
          { number: 4, maxPoints: 10, questionText: '', rating: 'Meets', score: 7, comments: '' },
          { number: 5, maxPoints: 10, questionText: '', rating: 'Meets', score: 8, comments: '' },
          { number: 6, maxPoints: 10, questionText: '', rating: 'Exceeds', score: 9, comments: '' }
        ],
        totalScore: 47,
        submittedAt: new Date().toISOString()
      })
    }
  };

  var result = doPost(testData);
  Logger.log(result.getContent());
}
