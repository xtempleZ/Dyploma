function main() {
  const docId = '13dgTkTQRzIEDc4bz_3SMkFR2AYlKkYR4uwEiAWhUGEM';
  const spreadsheetId = '1cdLPGCtZsCOL_QJC5gXsX5SNeQgX7bl2orLLJGNlKZM';
  const sheetName = '8ф-1к';

  const doc = DocumentApp.openById(docId);
  const text = doc.getBody().getText();

  const { date, names, unmatched: invalidNames } = parseAttendanceData(text);

  if (!date || (names.length === 0 && invalidNames.length === 0)) {
    Logger.log("Ошибка: не удалось распознать дату или корректные ФИО.");
    writeUnmatchedToDoc(doc, invalidNames.length > 0 ? invalidNames : ["(не удалось прочитать имена)"]);
    return;
  }

  const { unmatched, matched } = markAttendance(spreadsheetId, sheetName, date, names);

  writeMatchedToDoc(doc, matched);
  writeUnmatchedToDoc(doc, [...invalidNames, ...unmatched]);
}


// Извлекает дату и список ФИО из текста документа
function parseAttendanceData(text) {
  const lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
  const dateMatch = text.match(/\d{2}\.\d{2}\.\d{2}/);
  const date = dateMatch ? dateMatch[0] : null;

  const namePattern = /^[А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+$/;
  const capsPattern = /^[А-ЯЁ]+ [А-ЯЁ]+$/;
  const latinPattern = /[A-Za-z]/;

  const names = [];
  const unmatched = [];

  for (const line of lines) {
    if (latinPattern.test(line)) {
      unmatched.push(line);
    } else if (namePattern.test(line)) {
      names.push(line);
    } else if (capsPattern.test(line)) {
      // Преобразуем к нормальному виду (ИВАНОВ ПЕТР → Иванов Петр)
      const [last, first] = line.split(" ");
      const proper = capitalize(last) + " " + capitalize(first);
      names.push(proper);
    }
  }

  return { date, names, unmatched };
}

function capitalize(word) {
  if (!word) return "";
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}


// Сравнение с допущением неточностей (≥70%)
function findBestMatch(target, candidates) {
  const normalize = s => s.toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, '');
  const threshold = 0.7;

  function similarity(s1, s2) {
    const len = Math.max(s1.length, s2.length);
    let matches = 0;
    for (let i = 0; i < Math.min(s1.length, s2.length); i++) {
      if (s1[i] === s2[i]) matches++;
    }
    return matches / len;
  }

  const targetNorm = normalize(target);
  let bestMatch = null;
  let bestScore = 0;

  for (const candidate of candidates) {
    const score = similarity(targetNorm, normalize(candidate));
    if (score > bestScore) {
      bestScore = score;
      bestMatch = candidate;
    }
  }

  return bestScore >= threshold ? bestMatch : null;
}

// Отметка посещения в таблице и возврат нераспознанных
function markAttendance(sheetId, sheetName, dateStr, recognizedNames) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  const headerRowIndex = 2;
  const dateStartColIndex = 5;
  const nameStartRowIndex = 3;
  const nameColIndex = 3;

  const dateRow = data[headerRowIndex];
  let colIndex = -1;
  for (let i = dateStartColIndex; i < dateRow.length; i++) {
    const cellStr = formatDate(dateRow[i]);
    if (cellStr === dateStr) {
      colIndex = i;
      break;
    }
  }

  if (colIndex === -1) {
    Logger.log(`Дата ${dateStr} не найдена`);
    return { unmatched: recognizedNames, matched: [] };
  }

  const allNames = data.slice(nameStartRowIndex).map(row => row[nameColIndex]);
  const unmatched = [];
  const matched = [];

  for (const recogName of recognizedNames) {
    let found = false;

    for (let i = 0; i < allNames.length; i++) {
      const sheetName = allNames[i];
      if (!sheetName) continue;

      const match = findBestMatch(recogName, [sheetName]);
      if (match) {
        const rowToMark = nameStartRowIndex + i + 1;
        const colToMark = colIndex + 1;
        sheet.getRange(rowToMark, colToMark).setValue("•");
        matched.push(sheetName);
        found = true;
        break;
      }
    }

    if (!found) {
      unmatched.push(recogName);
    }
  }

  return { unmatched, matched };
}

// Форматирует дату из ячейки в формат "dd.mm.yy"
function formatDate(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd.MM.yy');
  }
  const str = cell.toString().trim();
  const match = str.match(/\d{2}\.\d{2}\.\d{2}/);
  return match ? match[0] : '';
}

// Запись нераспознанных ФИО в конец документа
function writeUnmatchedToDoc(doc, unmatchedList) {
  if (unmatchedList.length === 0) return;

  const body = doc.getBody();
  body.appendParagraph('\n\nНеопознанные имена:').setBold(true);
  unmatchedList.forEach(name => body.appendParagraph(name));
}

function writeMatchedToDoc(doc, matchedList) {
  const body = doc.getBody();
  body.appendParagraph('\n\nОтметка проставлена студентам:').setBold(true);
  if (matchedList.length === 0) {
    body.appendParagraph('Никому не удалось проставить отметку.');
  } else {
    matchedList.forEach(name => body.appendParagraph(name));
  }
}
