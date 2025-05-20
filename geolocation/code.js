const SHEET_ID = "1zhOA8GpU6a7GLauj8O0FJWP6zdF8CCe87L_QtYXUvJ8";
const REGISTRY_SHEET = "Реестр";
const ATTENDANCE_SHEET = "8ф-1к";

// База аудиторий с координатами
const CLASSROOMS = {
  "707А": {
    lat: 56.3427,
    lng: 38.2410,
    maxDistance: 100
  },
  "505А": {
    lat: 59.3435,
    lng: 39.2422,
    maxDistance: 100
  },
  "609Б": {
    lat: 50.3435,
    lng: 29.2422,
    maxDistance: 100
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Регистрация посещения")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Функция для получения списка аудиторий (для HTML)
function getClassroomsList() {
  return Object.keys(CLASSROOMS);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// Основная функция регистрации (обновленная)
function submitAttendance(lat, lng, acc, classroom) {
  const log = [];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const email = getUserEmail();
  log.push("Email: " + email);

  if (!email) throw new Error("Не удалось получить email пользователя.");
  
  // Проверка аудитории
  if (!CLASSROOMS[classroom]) {
    throw new Error("Неизвестная аудитория: " + classroom);
  }
  
  // Проверка местоположения
  const { lat: classLat, lng: classLng, maxDistance } = CLASSROOMS[classroom];
  const distance = getDistanceFromLatLng(lat, lng, classLat, classLng);
  log.push("Расстояние до аудитории: " + Math.round(distance) + " м");

  if (distance > maxDistance) {
    throw new Error(`Вы находитесь слишком далеко от ${classroom}. Максимальное расстояние: ${maxDistance} м`);
  }

  // Поиск пользователя в реестре
  const registrySheet = ss.getSheetByName(REGISTRY_SHEET);
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow()-1, 3).getValues();
  
  let fullName = null;
  for (let i = 0; i < registryData.length; i++) {
    if (registryData[i][1] === email) {
      fullName = registryData[i][2];
      break;
    }
  }
  
  if (!fullName) throw new Error("Email не найден в реестре.");
  log.push("ФИО найдено: " + fullName);

  // Отметка посещения
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");

  // Поиск столбца с сегодняшней датой
  const datesRow = sheet.getRange(3, 5, 1, sheet.getLastColumn()-4).getValues()[0];
  let dateCol = -1;
  
  for (let i = 0; i < datesRow.length; i++) {
    const cellDate = datesRow[i];
    if (cellDate instanceof Date) {
      const formattedDate = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "dd.MM.yyyy");
      if (formattedDate === today) {
        dateCol = i + 5;
        break;
      }
    }
  }
  
  if (dateCol === -1) throw new Error("Текущая дата не найдена в таблице посещений.");
  log.push("Найден столбец с датой: " + dateCol);

  // Поиск строки с ФИО
  const namesData = sheet.getRange(4, 4, sheet.getLastRow()-3, 1).getValues();
  let nameRow = -1;
  
  for (let i = 0; i < namesData.length; i++) {
    if (namesData[i][0] === fullName) {
      nameRow = i + 4;
      break;
    }
  }
  
  if (nameRow === -1) throw new Error("ФИО не найдено на листе посещаемости.");
  log.push("Найдено ФИО в строке: " + nameRow);

  // Ставим отметку
  sheet.getRange(nameRow, dateCol).setValue("·");

  return {
    status: "ok",
    message: `Посещение аудитории ${classroom} успешно зарегистрировано для ${fullName}`,
    log: log
  };
}

function getDistanceFromLatLng(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}
