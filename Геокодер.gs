// Конфигурация
var DADATA_API_KEY = "afd212549d135d9b8bdd1a607ed35e9a4d48c601";
var DADATA_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs";

// Главная функция для показа карты
function showPostcardMap() {
  try {
    // Создаем HTML с картой
    var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(1500)
      .setHeight(1000);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Карта отправленных открыток');
    
  } catch (error) {
    Logger.log('Ошибка в showPostcardMap: ' + error.toString());
    // Показываем ошибку
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка загрузки карты: ' + error.message);
  }
}

function showDeskView() {
  try {

    syncNewSourcesFromSheet();

    var html = HtmlService.createHtmlOutputFromFile('desk')
      .setWidth(1500)
      .setHeight(1700);
    
    SpreadsheetApp.getUi().showModalDialog(html, ' ');
    
  } catch (error) {
    Logger.log('Ошибка в showDeskView: ' + error.toString());
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка загрузки стола: ' + error.message);
  }
}

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📮 Открытки')
      .addItem('📍 Показать карту', 'showPostcardMap')
      .addItem('✉️ Показать стол', 'showDeskView')
      .addToUi();
    
    // Автоматически синхронизируем источники при открытии
    syncSourcesWithSheet();
    
  } catch (error) {
    Logger.log('Ошибка создания меню: ' + error.toString());
  }
}


// Функция для получения данных (используется в index.html)
function getPostcardsDataForHTML() {
  return getPostcardsData();
}

// Функция для получения данных из таблицы
function getPostcardsData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    var postcards = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Проверяем столбец с координатами (индекс 11)
      if (row[11] && row[11].toString().includes(',')) {
        var postcard = {
          id: row[2] || '',
          source: row[1] || '',
          status: row[6] || '',
          daysInTransit: row[5] || '',
          name: row[7] || '',
          country: row[8] || '',
          city: row[9] || '',
          index: row[10] || '',
          coordinates: row[11] || '',
          sentDate: row[3] ? formatDate(row[3]) : '',
          receivedDate: row[4] ? formatDate(row[4]) : ''
        };
        postcards.push(postcard);
      }
    }
    
    Logger.log('Загружено открыток: ' + postcards.length);
    return postcards;
  } catch (error) {
    Logger.log('Ошибка получения данных: ' + error.toString());
    return [];
  }
}

// Функция-обработчик выбора в меню
function onMenuSelect(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // Проверяем, что изменение в нужном листе и ячейке
    if (sheet.getName() === 'Статистика' && range.getA1Notation() === 'M1') {
      const selectedValue = e.value;
      
      Logger.log('Выбрано в меню: ' + selectedValue);
      
      if (selectedValue === 'Открыть карту') {
        openMapFromMenu();
      } else if (selectedValue === 'Открыть стол') {
        openDeskFromMenu();
      }
      
      // Очищаем ячейку после выполнения
      SpreadsheetApp.flush();
      range.clear();
      Logger.log('Ячейка M1 очищена');
    }
  } catch (error) {
    Logger.log('Ошибка в onMenuSelect: ' + error.toString());
  }
}

// Функция для открытия карты из меню
function openMapFromMenu() {
  try {
    // Просто вызываем существующую функцию показа карты
    showPostcardMap();
    Logger.log('Карта открыта через меню');
  } catch (error) {
    Logger.log('Ошибка открытия карты из меню: ' + error.toString());
    
    // Показываем ошибку пользователю
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка открытия карты: ' + error.message);
  }
}

// Функция для открытия стола из меню
function openDeskFromMenu() {
  try {
    // Просто вызываем существующую функцию показа стола
    showDeskView();
    Logger.log('Стол открыт через меню');
  } catch (error) {
    Logger.log('Ошибка открытия стола из меню: ' + error.toString());
    
    // Показываем ошибку пользователю
    var ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка открытия стола: ' + error.message);
  }
}

function getStatusCounts() {
  try {
    // Синхронизируем источники при подсчете статусов
    syncNewSourcesFromSheet();
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    var statusCounts = {
      'В процессе': 0,
      'Готово к отправке': 0,
      'Нет открытки': 0,
      'Нет марки': 0,
      'Карта': 0
    };
    
    var travelStatuses = ['В пути', 'Потеряно', 'Получено'];
    var travelCount = 0;
    
    // Считаем все за один проход
   for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[6] ? row[6].toString().trim() : '';
      
      if (statusCounts.hasOwnProperty(status)) {
        statusCounts[status]++;
      }
      
      if (travelStatuses.includes(status)) {
        travelCount++;
      }
    }
    
    statusCounts['Карта'] = travelCount;
    
    Logger.log('Посчитаны статусы: ' + JSON.stringify(statusCounts));
    return statusCounts;
    
  } catch (error) {
    Logger.log('Ошибка подсчета статусов: ' + error.toString());
    return {
      'В процессе': 0,
      'Готово к отправке': 0,
      'Нет открытки': 0,
      'Нет марки': 0,
      'Карта': 0
    };
  }
}

// Функция для получения списка открыток по статусу (для стола)
function getCardsByStatus(status) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    var cards = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cardStatus = row[6] || '';
      
      if (cardStatus === status) {
        var card = {
          id: row[2] || '',
          name: row[7] || 'Не указано',
          country: row[8] || 'Не указано',
          source: row[1] || 'Не указано',
          city: row[9] || 'Не указано',
          index: row[10] || 'Не указано',
          sentDate: row[3] ? formatDate(row[3]) : 'Не указано',
          daysInTransit: row[5] || 'Не указано',
          receivedDate: row[4] ? formatDate(row[4]) : 'Не указано'
        };
        cards.push(card);
      }
    }
    
    Logger.log('Найдено открыток со статусом "' + status + '": ' + cards.length);
    return cards;
  } catch (error) {
    Logger.log('Ошибка получения карточек: ' + error.toString());
    return [];
  }
}

function getStatusImages() {
  var imageMap = {
    'Нет открытки': 'https://ibb.co/N2BVcnMr/image.jpg',
    'В процессе': 'https://i.ibb.co/qYQyTFvC',
    'Нет марки': 'https://i.ibb.co/qYQyTFvC',
    'Готово к отправке': 'https://i.ibb.co/qYQyTFvC'
  };

  return imageMap;
}


// Вспомогательная функция для форматирования дат в формате "дд.мм.гггг"
function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  
  // Если дата пришла как строка в формате "гггг-мм-дд", преобразуем её
  if (typeof date === 'string' && date.match(/^\d{4}-\d{2}-\d{2}$/)) {
    try {
      var parts = date.split('-');
      var year = parts[0];
      var month = parts[1];
      var day = parts[2];
      return day + '.' + month + '.' + year;
    } catch (e) {
      return date; // Возвращаем как есть при ошибке
    }
  }
  
  return date;
}


// Функция для получения всех открыток
function getAllCards() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    var cards = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var card = {
        id: row[2] || '',
        source: row[1] || '',
        status: row[6] || '',
        daysInTransit: row[5] || '',
        name: row[7] || '',
        country: row[8] || '',
        city: row[9] || '',
        index: row[10] || '',
        coordinates: row[11] || '',
        sentDate: row[3] ? formatDate(row[3]) : '',
        receivedDate: row[4] ? formatDate(row[4]) : ''
      };
      cards.push(card);
    }
    
    Logger.log('Загружено всех открыток: ' + cards.length);
    return cards;
  } catch (error) {
    Logger.log('Ошибка получения всех открыток: ' + error.toString());
    return [];
  }
}

// Функция для изменения статуса открытки
function updateCardStatus(cardId, newStatus) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    // Ищем открытку по ID в столбце C (индекс 2)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var currentCardId = row[2] ? row[2].toString().trim() : '';
      
      if (currentCardId === cardId) {
        // Обновляем статус в столбце G (индекс 6)
        sheet.getRange(i + 1, 7).setValue(newStatus);
        Logger.log('Статус открытки ' + cardId + ' изменен на: ' + newStatus);
        return { success: true, message: 'Статус обновлен!' };
      }
    }
    
    return { success: false, message: 'Открытка с ID ' + cardId + ' не найдена' };
    
  } catch (error) {
    Logger.log('Ошибка обновления статуса: ' + error.toString());
    return { success: false, message: 'Ошибка: ' + error.message };
  }
}

// Функция для смены статуса через выпадающий список
function showStatusChangeDialog(cardId, currentStatus) {
  try {
    // Создаем выпадающий список прямо в ячейке
    var statusOptions = ['В процессе', 'Готово к отправке', 'Нет открытки', 'Нет марки'];
    
    // Просто обновляем статус без диалогового окна
    return { 
      success: true, 
      statusOptions: statusOptions,
      currentStatus: currentStatus
    };
    
  } catch (error) {
    Logger.log('Ошибка в showStatusChangeDialog: ' + error.toString());
    return { success: false, message: error.message };
  }
}

// Функция для быстрого обновления статуса без перезагрузки страницы
function quickUpdateStatus(cardId, newStatus) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    // Ищем открытку по ID в столбце C (индекс 2)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var currentCardId = row[2] ? row[2].toString().trim() : '';
      
      if (currentCardId === cardId) {
        // Обновляем статус в столбце G (индекс 6)
        sheet.getRange(i + 1, 7).setValue(newStatus);
        Logger.log('Статус открытки ' + cardId + ' изменен на: ' + newStatus);
        
        // Если статус меняется на "Получено", можно автоматически проставить дату получения
        if (newStatus === 'Получено') {
          var today = new Date();
          sheet.getRange(i + 1, 5).setValue(today); // Столбец E - дата получения
          Logger.log('Автоматически проставлена дата получения для открытки: ' + cardId);
        }
        
        return { 
          success: true, 
          message: 'Статус обновлен на: ' + newStatus,
          newStatus: newStatus,
          cardId: cardId
        };
      }
    }
    
    return { success: false, message: 'Открытка с ID ' + cardId + ' не найдена' };
    
  } catch (error) {
    Logger.log('Ошибка обновления статуса: ' + error.toString());
    return { success: false, message: 'Ошибка: ' + error.message };
  }
}

// Функция для получения открыток со статусами доставки
function getTravelStatusCards() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    var cards = [];
    var travelStatuses = ['В пути', 'Потеряно', 'Получено'];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cardStatus = row[6] || '';
      
      // Фильтруем только статусы доставки
      if (travelStatuses.includes(cardStatus)) {
        var card = {
          id: row[2] || '',
          source: row[1] || '',
          status: cardStatus,
          daysInTransit: row[5] || '',
          name: row[7] || '',
          country: row[8] || '',
          city: row[9] || '',
          index: row[10] || '',
          coordinates: row[11] || '',
          sentDate: row[3] ? formatDate(row[3]) : '',
          receivedDate: row[4] ? formatDate(row[4]) : ''
        };
        cards.push(card);
      }
    }
    
    Logger.log('Найдено открыток со статусами доставки: ' + cards.length);
    return cards;
  } catch (error) {
    Logger.log('Ошибка получения карточек доставки: ' + error.toString());
    return [];
  }
}

// Функция для получения всех данных конкретной открытки
function getCardData(cardId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var currentCardId = row[2] ? row[2].toString().trim() : '';
      
      if (currentCardId === cardId) {
        var card = {
          id: cardId,
          // Все поля из таблицы
          source: row[1] || '',
          sentDate: row[3] ? formatDateForEdit(row[3]) : '',
          receivedDate: row[4] ? formatDateForEdit(row[4]) : '',
          daysInTransit: row[5] || '',
          status: row[6] || '',
          name: row[7] || '',
          country: row[8] || '',
          city: row[9] || '',
          index: row[10] || '',
          coordinates: row[11] || '',
          notes: row[12] || '', // если есть поле с заметками
          imageUrl: row[13] || '' // если есть поле с изображением
        };
        return { success: true, card: card };
      }
    }
    
    return { success: false, message: 'Открытка не найдена' };
    
  } catch (error) {
    Logger.log('Ошибка получения данных открытки: ' + error.toString());
    return { success: false, message: 'Ошибка: ' + error.message };
  }
}

// Функция для обновления всех данных открытки
function updateCardData(cardId, updatedData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var currentCardId = row[2] ? row[2].toString().trim() : '';
      
      if (currentCardId === cardId) {
        // Обновляем все поля в соответствующих столбцах
        // Источник - столбец B (индекс 1)
        sheet.getRange(i + 1, 2).setValue(updatedData.source || '');
        // Дата отправки - столбец D (индекс 3)
        sheet.getRange(i + 1, 4).setValue(updatedData.sentDate || '');
        // Дата получения - столбец E (индекс 4)
        sheet.getRange(i + 1, 5).setValue(updatedData.receivedDate || '');
        // Дней в пути - столбец F (индекс 5)
        sheet.getRange(i + 1, 6).setValue(updatedData.daysInTransit || '');
        // Статус - столбец G (индекс 6)
        sheet.getRange(i + 1, 7).setValue(updatedData.status || '');
        // Имя - столбец H (индекс 7)
        sheet.getRange(i + 1, 8).setValue(updatedData.name || '');
        // Страна - столбец I (индекс 8)
        sheet.getRange(i + 1, 9).setValue(updatedData.country || '');
        // Город - столбец J (индекс 9)
        sheet.getRange(i + 1, 10).setValue(updatedData.city || '');
        // Индекс - столбец K (индекс 10)
        sheet.getRange(i + 1, 11).setValue(updatedData.index || '');
        // Координаты - столбец L (индекс 11)
        sheet.getRange(i + 1, 12).setValue(updatedData.coordinates || '');
        // Заметки - столбец M (индекс 12)
        sheet.getRange(i + 1, 13).setValue(updatedData.notes || '');
        // Изображение - столбец N (индекс 13)
        sheet.getRange(i + 1, 14).setValue(updatedData.imageUrl || '');
        
        Logger.log('Все данные открытки ' + cardId + ' обновлены');
        return { success: true, message: 'Данные сохранены' };
      }
    }
    
    return { success: false, message: 'Открытка с ID ' + cardId + ' не найдена' };
    
  } catch (error) {
    Logger.log('Ошибка обновления данных открытки: ' + error.toString());
    return { success: false, message: 'Ошибка: ' + error.message };
  }
}

// Функция для форматирования даты для формы редактирования
function formatDateForEdit(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return date;
}

// ИСТОЧНИКИ //

// Основной список источников (хранится в коде + PropertiesService)
const DEFAULT_SOURCES = [
  'Postcrossing',
  'Домоткрыток', 
  'PostFun',
  'Личные',
  'Другое'
];

// Получает источники для интерфейса
function getSources() {
  try {
    // Пытаемся получить из PropertiesService
    var savedSources = PropertiesService.getScriptProperties().getProperty('SOURCES');
    if (savedSources) {
      var parsedSources = JSON.parse(savedSources);
      if (parsedSources && parsedSources.length > 0) {
        return parsedSources;
      }
    }
    
    // Если нет в PropertiesService, используем дефолтные
    // И сохраняем их для будущего использования
    PropertiesService.getScriptProperties()
      .setProperty('SOURCES', JSON.stringify(DEFAULT_SOURCES));
    
    // Синхронизируем с таблицей
    syncSourcesWithSheet();
    
    return DEFAULT_SOURCES;
    
  } catch (error) {
    Logger.log('Ошибка получения источников: ' + error.toString());
    return DEFAULT_SOURCES;
  }
}

// Добавить новый источник
function addNewSource(newSource) {
  if (!newSource || newSource.trim() === '') {
    return { success: false, message: 'Название источника не может быть пустым' };
  }
  
  var source = newSource.trim();
  
  // Получаем текущие источники
  var currentSources = getSources();
  
  // Проверяем, нет ли уже такого источника
  if (currentSources.includes(source)) {
    return { success: false, message: 'Этот источник уже существует' };
  }
  
  // Добавляем в список и сортируем
  currentSources.push(source);
  currentSources.sort();
  
  // Сохраняем в PropertiesService
  PropertiesService.getScriptProperties()
    .setProperty('SOURCES', JSON.stringify(currentSources));
  
  // Синхронизируем с таблицей
  syncSourcesWithSheet();
  
  return { 
    success: true, 
    message: 'Источник "' + source + '" добавлен',
    sources: currentSources 
  };
}

// Синхронизация источника с выпадающим списком в таблице "От меня"
function syncSourcesWithSheet() {
  try {
    var sources = getSources();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    
    var validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(sources, true)
      .setAllowInvalid(true)
      .setHelpText('Выберите источник из списка')
      .build();
    
    // Применяем к столбцу B (источники)
    sheet.getRange('B2:B').setDataValidation(validation);
    
    Logger.log('Синхронизированы источники: ' + sources.join(', '));
    
  } catch (error) {
    Logger.log('Ошибка синхронизации источников: ' + error.toString());
  }
}

// Синхронизация с защитой от частых вызовов
function syncNewSourcesFromSheet() {
  // Проверяем, когда последний раз синхронизировали (не чаще чем раз в 5 минут)
  var lastSync = PropertiesService.getScriptProperties().getProperty('LAST_SOURCES_SYNC');
  var now = new Date().getTime();
  
  if (lastSync && (now - parseInt(lastSync)) < 5 * 60 * 1000) { // 5 минут
    Logger.log('Синхронизация пропущена (слишком частая)');
    return { success: true, added: [], skipped: true };
  }
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var data = sheet.getRange('B2:B').getValues();
    
    var currentSources = getSources();
    var sourcesInSheet = new Set();
    
    // Собираем ВСЕ источники, используемые в таблице
    for (var i = 0; i < data.length; i++) {
      var source = data[i][0];
      if (source && source.toString().trim() !== '') {
        sourcesInSheet.add(source.toString().trim());
      }
    }
    
    // Находим новые источники
    var newSources = Array.from(sourcesInSheet).filter(source => 
      !currentSources.includes(source)
    );
    
    if (newSources.length > 0) {
      Logger.log('Найдены новые источники в таблице: ' + newSources.join(', '));
      
      var updatedSources = currentSources.concat(newSources).sort();
      
      // Сохраняем обновленный список
      PropertiesService.getScriptProperties()
        .setProperty('SOURCES', JSON.stringify(updatedSources));
      
      // Сохраняем время последней синхронизации
      PropertiesService.getScriptProperties()
        .setProperty('LAST_SOURCES_SYNC', now.toString());
      
      // Синхронизируем валидацию
      syncSourcesWithSheet();
      
      Logger.log('Добавлены новые источники: ' + newSources.join(', '));
      return { success: true, added: newSources, allSources: updatedSources };
    }
    
    // Все равно сохраняем время синхронизации
    PropertiesService.getScriptProperties()
      .setProperty('LAST_SOURCES_SYNC', now.toString());
    
    return { success: true, added: [], allSources: currentSources };
    
  } catch (error) {
    Logger.log('Ошибка синхронизации источников из таблицы: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Добавление открытки
function addNewPostcardToSheet(cardData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('От меня');
    var lastRow = sheet.getLastRow();
    
    // Автогенерируем ID
    var newId = 'PC' + (new Date().getTime()).toString().slice(-6);
    
    // Добавляем новую строку
    var newRow = [
      '', // A - пусто
      cardData.source || 'Не указано', // B - источник
      newId, // C - ID
      '', // D - дата отправки
      '', // E - дата получения  
      '', // F - дней в пути
      cardData.status || 'В процессе', // G - статус
      cardData.name || '', // H - имя
      cardData.country || '', // I - страна
      cardData.city || '', // J - город
      cardData.index || '', // K - индекс
      '', // L - координаты
      cardData.notes || '', // M - заметки
      cardData.imageUrl || ''  // N - изображение
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    return { success: true, message: 'Открытка добавлена с ID: ' + newId, cardId: newId };
    
  } catch (error) {
    Logger.log('Ошибка добавления открытки: ' + error.toString());
    return { success: false, message: 'Ошибка: ' + error.toString() };
  }
}

/**
 * Найти почтовое отделение по индексу
 * @param {string} index Почтовый индекс
 * @param {string} returnType Что вернуть: "address", "coords", "full"
 * @return {string} Запрошенная информация
 * @customfunction
 */
function GET_POST_OFFICE(index, returnType = "address") {
  if (!index) return "Введите индекс";
  
  var url = DADATA_URL + "/findById/postal_unit";
  var payload = { "query": index.toString() };
  
  var options = {
    'method': 'POST',
    'headers': {
      'Authorization': 'Token ' + DADATA_API_KEY,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    
    if (data.suggestions && data.suggestions.length > 0) {
      var office = data.suggestions[0].data;
      
      switch(returnType.toLowerCase()) {
        case "address":
          return office.address_str || data.suggestions[0].value;
        
        case "coords":
          if (office.geo_lat && office.geo_lon) {
            return office.geo_lat + "," + office.geo_lon;
          }
          return "Координаты не указаны";
        
        case "full":
          return getFullInfo(data.suggestions[0]);
        
        default:
          return office.address_str || data.suggestions[0].value;
      }
    }
    return "Отделение не найдено";
  } catch (e) {
    return "Ошибка: " + e.toString();
  }
}

/**
 * Поиск отделений по адресу
 * @param {string} query Адрес или часть адреса
 * @param {number} count Количество результатов (1-20)
 * @return {string} Адреса отделений
 * @customfunction
 */
function SEARCH_POST_OFFICES(query, count = 5) {
  if (!query) return "Введите запрос";
  
  var url = DADATA_URL + "/suggest/postal_unit";
  var payload = {
    "query": query.toString(),
    "count": Math.min(count, 20)
  };
  
  var options = {
    'method': 'POST',
    'headers': {
      'Authorization': 'Token ' + DADATA_API_KEY,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    
    if (data.suggestions && data.suggestions.length > 0) {
      var results = [];
      for (var i = 0; i < data.suggestions.length; i++) {
        var office = data.suggestions[i];
        var status = office.data.is_closed ? " (ЗАКРЫТО)" : "";
        results.push((i+1) + ". " + office.value + status);
      }
      return results.join("\n");
    }
    return "Отделения не найдены";
  } catch (e) {
    return "Ошибка: " + e.toString();
  }
}

/**
 * Найти ближайшие отделения по координатам
 * @param {number} lat Широта
 * @param {number} lon Долгота
 * @param {number} radius Радиус поиска в метрах (по умолчанию 1000)
 * @return {string} Список ближайших отделений
 * @customfunction
 */
function NEAREST_POST_OFFICES(lat, lon, radius = 1000) {
  if (!lat || !lon) return "Введите координаты";
  
  var url = DADATA_URL + "/geolocate/postal_unit";
  var payload = {
    "lat": parseFloat(lat),
    "lon": parseFloat(lon),
    "radius_meters": parseInt(radius)
  };
  
  var options = {
    'method': 'POST',
    'headers': {
      'Authorization': 'Token ' + DADATA_API_KEY,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    
    if (data.suggestions && data.suggestions.length > 0) {
      var results = [];
      for (var i = 0; i < Math.min(data.suggestions.length, 5); i++) {
        var office = data.suggestions[i];
        var distance = office.distance ? " (" + Math.round(office.distance) + "м)" : "";
        var status = office.data.is_closed ? " - ЗАКРЫТО" : "";
        results.push((i+1) + ". " + office.value + distance + status);
      }
      return results.join("\n");
    }
    return "Ближайшие отделения не найдены";
  } catch (e) {
    return "Ошибка: " + e.toString();
  }
}

function getFullInfo(suggestion) {
  var office = suggestion.data;
  var info = [
    "АДРЕС: " + suggestion.value,
    "ИНДЕКС: " + office.postal_code,
    "КООРДИНАТЫ: " + (office.geo_lat && office.geo_lon ? office.geo_lat + "," + office.geo_lon : "не указаны"),
  ];
  
  return info.join("\n");
}

// Функция для принудительного создания меню
function initMenu() {
  onOpen();
}
