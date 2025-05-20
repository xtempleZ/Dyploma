// ID таблицы с расписанием
var IDTT = '1eBTAl0uf-cNvrWpgj5MtjNyo8favahbM1y_UfpCbYV0'; // Замените на ваш ID

// Функция для добавления меню в Google Таблицы
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Запустить перенос', 'runAddon')
    .addToUi();
  runAddon();
}

// Функция для установки
function onInstall() {
  onOpen();
  tekWeek();
}

// Функция для запуска аддона
function runAddon() {
  var ui = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Перенос расписания')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(ui);
}

// Функция для получения данных о текущем семестре
function TekSemestr() {
  var STT = SpreadsheetApp.openById(IDTT).getSheetByName('таблицы');
  var numTek = STT.getRange('A2').getValue();
  var Ids = STT.getRange('B' + numTek + ':G' + numTek).getValues();
  return Ids[0];
}

// Функция для преобразования времени в миллисекунды
function constHOURS(fromStr) {
  let sepIdx1 = fromStr.indexOf(":");
  let hoursStr = fromStr.substring(0, sepIdx1);
  let minsStr = fromStr.substring(sepIdx1 + 1);
  return 1000 * 60 * 60 * hoursStr + 1000 * 60 * minsStr;
}

// Функция для удаления всех событий из календаря
function removeCalendar() {
  var cal = CalendarApp.getCalendarById("8bfb223f727fc1459b06f52c750c0464b1c7bee6b609a1ffc0b2a8768c3f50de@group.calendar.google.com"); // Замените на ваш календарь
  var now = new Date();
  var twoHoursFromNow = new Date(now.getTime() + (5 * 31 * 24 * 60 * 60 * 1000));
  var events = cal.getEvents(now, twoHoursFromNow);
  events.forEach(e => e.deleteEvent());
  Logger.log("Все события удалены.");
}

function tekWeek(d){

  if (!d) d=new Date();
  else d=new Date(d);

  if (d.getDay()==0) d.setDate(d.getDate() + 1);
  // let week=d.getWeek();
  let tabl=SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getDataRange().getValues();
  let n=2;
  for ( let i=2; i<tabl.length;i++){
   n=i;
  //  let week1=tabl[i][2].getWeek();
   if(tabl[i][2]<=d && d<=tabl[i][3] ) break;
 //  if(week==week1 ) break;
  }
  let s=SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  let r=s.getRange(n+1,2,1,4);
  s.setActiveSelection(r);
  let dn=d.getDay();
  let vn=r.getValues()[0][3];
  let j=0;
  if (vn=="нижняя") j=2*dn+3;
  else j=2*dn+2;

  let s0=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  // let r1=s0.getRange(j,1,1,1);
  let r1=s0.getRange(j+':'+j);
  s0.setActiveSelection(r1);
  return n;
}

function createTopFoot(dateStr, isTopWeek) {
  try {
    // Проверяем, что дата передана
    if (!dateStr) {
      throw new Error("Дата начала семестра не указана.");
    }

    // Преобразуем строку даты в объект Date
    var startDate = new Date(dateStr);
    if (isNaN(startDate.getTime())) {
      throw new Error("Некорректная дата начала семестра.");
    }

    // Устанавливаем время на 00:00:00
    startDate.setHours(0, 0, 0, 0);

    // Получаем лист "верх/низ"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('верх/низ');
    if (!sheet) {
      throw new Error("Лист 'верх/низ' не найден.");
    }

    // Очищаем старые данные (если нужно)
    sheet.getRange("B3:E20").clearContent().clearFormat();

    // Создаем массив для данных
    var data = [];
    var backgrounds = [];

    // Начинаем с выбранной даты
    var currentDate = new Date(startDate);
    var weekType = isTopWeek ? "верхняя" : "нижняя";

    // Создаем 18 недель (как в вашем коде)
    for (var i = 0; i < 18; i++) {
      // Начало недели (понедельник)
      var weekStart = new Date(currentDate);
      weekStart.setDate(currentDate.getDate() - currentDate.getDay() + 1);

      // Конец недели (воскресенье)
      var weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);

      // Добавляем данные в массив
      data.push([i + 1, weekStart, weekEnd, weekType]);

      // Добавляем цвет фона
      var bgColor = weekType === "верхняя" ? "#d9ead3" : "#bf9000";
      backgrounds.push([bgColor, bgColor, bgColor, bgColor]);

      // Переключаем тип недели
      weekType = weekType === "верхняя" ? "нижняя" : "верхняя";

      // Переходим к следующей неделе
      currentDate.setDate(currentDate.getDate() + 7);
    }

    // Записываем данные в таблицу
    sheet.getRange(3, 2, 18, 4).setValues(data).setBackgrounds(backgrounds);

    Logger.log("Таблица недель успешно создана.");
  } catch (e) {
    Logger.log("Ошибка при создании таблицы недель: " + e.toString());
  }
}

// Основная функция для переноса расписания в календарь
function calendar() {
  try {
    var cal = CalendarApp.getCalendarById("8bfb223f727fc1459b06f52c750c0464b1c7bee6b609a1ffc0b2a8768c3f50de@group.calendar.google.com"); // Замените на ваш календарь
    if (!cal) {
      throw new Error("Календарь не найден или недоступен.");
    }

    var spreadsheet = SpreadsheetApp.getActive();
    var teklist = spreadsheet.getSheetByName('верх/низ'); // Лист с датами верхних/нижних недель
    var teklist2 = spreadsheet.getSheetByName('2024 ВЕСНА'); // Лист с расписанием

    if (!teklist || !teklist2) {
      throw new Error("Листы 'верх/низ' или '2024 ВЕСНА' не найдены.");
    }

    // Получаем данные из листов
    var data = teklist.getRange("A3:E" + teklist.getLastRow()).getValues();
    var data2 = teklist2.getRange("A1:I" + teklist2.getLastRow()).getValues();

    Logger.log("Данные из 'верх/низ': " + JSON.stringify(data));
    Logger.log("Данные из '2024 ВЕСНА': " + JSON.stringify(data2));

    var y = 0; // Счетчик для верхних недель
    var t = 0; // Счетчик для нижних недель

    // Обрабатываем каждую строку расписания
    for (var j = 3; j < data2.length; j++) {
      if (j % 2 == 1) { // Верхняя неделя
        var couple = '';
        y += 1;
        for (var column = 2; column < data2[0].length; column++) {
          couple = data2[j][column];
          var row = j + 1;
          var data23 = teklist2.getRange(row, column + 1);
          if (couple == '' && data23.isPartOfMerge()) {
            couple = data23.getMergedRanges()[0].getCell(1, 1).getValue();
          }
          if (couple != '') {
            for (var i = 0; i < data.length; i++) {
              if (i % 2 == 1) { // Верхняя неделя
                var startDate = new Date(data[i][2]);
                const MILLIS_PER_DAY = 1000 * 60 * 60 * 24 * (y - 1);
                const newDate = new Date(startDate.getTime() + MILLIS_PER_DAY);
                newDate.setHours(0);

                var time = data2[2][column];
                let sepIdx = time.indexOf("-");
                let fromStr = time.substring(0, sepIdx).trim();
                let toStr = time.substring(sepIdx + 1).trim();

                var stime = new Date(newDate.getTime() + constHOURS(fromStr));
                var etime = new Date(newDate.getTime() + constHOURS(toStr));

                cal.createEvent(couple, stime, etime);
                Logger.log("Создано событие: " + couple + " с " + stime + " по " + etime);
              }
            }
          }
        }
      } else if (j % 2 == 0) { // Нижняя неделя
        var couple = '';
        t += 1;
        for (var column = 2; column < data2[0].length; column++) {
          couple = data2[j][column];
          var row = j + 1;
          var data23 = teklist2.getRange(row, column + 1);
          if (couple == '' && data23.isPartOfMerge()) {
            couple = data23.getMergedRanges()[0].getCell(1, 1).getValue();
          }
          if (couple != '') {
            for (var i = 0; i < data.length; i++) {
              if (i % 2 == 0) { // Нижняя неделя
                var startDate = new Date(data[i][2]);
                const MILLIS_PER_DAY = 1000 * 60 * 60 * 24 * (t - 1);
                const newDate = new Date(startDate.getTime() + MILLIS_PER_DAY);
                newDate.setHours(0);

                var time = data2[2][column];
                let sepIdx = time.indexOf("-");
                let fromStr = time.substring(0, sepIdx).trim();
                let toStr = time.substring(sepIdx + 1).trim();

                var stime = new Date(newDate.getTime() + constHOURS(fromStr));
                var etime = new Date(newDate.getTime() + constHOURS(toStr));

                cal.createEvent(couple, stime, etime);
                Logger.log("Создано событие: " + couple + " с " + stime + " по " + etime);
              }
            }
          }
        }
      }
    }
    Logger.log("Перенос расписания завершен.");
  } catch (e) {
    Logger.log("Ошибка: " + e.toString());
  }
}
