var SHEET_ID = '1HO_o7dbnc2GX_nfEFiiuVV9MWTOxD2IBepRcoH5cEGo';
var SHEET_NAME = '8ф-1к';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Учёт посещаемости')
    .addItem('Открыть панель', 'showSidebar')
    .addToUi();
  showSidebar();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Учёт посещаемости')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getStudentsAndDates() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    var lastRow = sheet.getLastRow();

    // ФИО: D4:D, Группа: E4:E
    var nameData = sheet.getRange(4, 4, lastRow - 3).getValues();  // D
    var groupData = sheet.getRange(4, 5, lastRow - 3).getValues(); // E

    var students = [];
    for (var i = 0; i < nameData.length; i++) {
      var name = nameData[i][0];
      var group = groupData[i][0];
      if (name && name.toString().trim() !== "") {
        students.push({ name: name, group: group });
      }
    }

    // Заголовки дат: строка 3, начиная с F3 (столбец 6) до "пропуски"
    var headerRow = sheet.getRange(3, 6, 1, sheet.getLastColumn() - 5).getValues()[0];

    var lastDateIndex = -1;
    for (var j = 0; j < headerRow.length; j++) {
      var cell = headerRow[j];
      if (typeof cell === "string" && cell.toLowerCase().indexOf("пропуск") !== -1) {
        lastDateIndex = j;
        break;
      }
    }
    if (lastDateIndex === -1) throw new Error("Не найден столбец 'пропуски'");

    var rawDates = headerRow.slice(0, lastDateIndex);
    var formattedDates = rawDates.map(function (d) {
      return Object.prototype.toString.call(d) === "[object Date]"
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd.MM.yyyy")
        : d.toString();
    });

    return {
      students: students,
      dates: formattedDates
    };

  } catch (e) {
    Logger.log("Ошибка в getStudentsAndDates: " + e.toString());
    throw e;
  }
}




function setMark(studentName, markValue, dateStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var names = sheet.getRange(4, 4, sheet.getLastRow() - 3).getValues(); // D4:D
    var dates = sheet.getRange(3, 5, 1, sheet.getLastColumn() - 4).getValues()[0]; // E3 →

    var rowIndex = -1;
    for (var i = 0; i < names.length; i++) {
      if (names[i][0] === studentName) {
        rowIndex = i + 4;
        break;
      }
    }

    var colIndex = -1;
    for (var j = 0; j < dates.length; j++) {
      var cellDate = dates[j];
      var dateFormatted = (cellDate instanceof Date)
        ? Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "dd.MM.yyyy")
        : cellDate;

      if (dateFormatted === dateStr) {
        colIndex = j + 5;
        break;
      }
    }

    if (rowIndex >= 0 && colIndex >= 0) {
      sheet.getRange(rowIndex, colIndex).setValue(markValue);
      return "✅ Поставлено значение '" + markValue + "' студенту " + studentName + " за " + dateStr;
    } else {
      throw new Error("ФИО или дата не найдены.");
    }
  } catch (e) {
    Logger.log("Ошибка в setMark: " + e.toString());
    throw e;
  }
}

function getMarksMap(dateStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();

    var nameRange = sheet.getRange(4, 4, lastRow - 3); // D4:D
    var nameRows = nameRange.getValues();
    var names = [];
    for (var i = 0; i < nameRows.length; i++) {
      names.push(nameRows[i][0]);
    }

    var dateRow = sheet.getRange(3, 6, 1, sheet.getLastColumn() - 5).getValues()[0];

    var colIndex = -1;
    for (var j = 0; j < dateRow.length; j++) {
      var cell = dateRow[j];
      var formatted = (cell instanceof Date)
        ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd.MM.yyyy")
        : cell;
      if (formatted === dateStr) {
        colIndex = j + 6; // потому что начинается с 6
        break;
      }
    }

    if (colIndex === -1) throw new Error("Дата не найдена в заголовке таблицы");

    var values = sheet.getRange(4, colIndex, lastRow - 3).getValues();

    var result = {};
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      if (name && name.toString().trim() !== "") {
        result[name] = values[i][0] ? values[i][0].toString() : "";
      }
    }

    return result;
  } catch (e) {
    Logger.log("Ошибка в getMarksMap: " + e.toString());
    throw e;
  }
}




