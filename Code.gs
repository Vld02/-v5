/*************************************************
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 *************************************************/
const WEBAPP_FAVICON_URL = 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.ico'; // Вставьте прямую HTTPS-ссылку на PNG/ICO для вкладки в обёртке Google Script.

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ',
  RESULT_SHEET_NAME: 'Результат',
  LOG_SHEET_NAME: 'Входы',
  YELLOW: '#ffff00',
  TIMEZONE: 'GMT+3',
  DATE_FORMAT: 'dd.MM.yyyy',
  NAMES_CACHE_KEY: 'dbv5_full_names_v1',
  NAMES_CACHE_TTL_SECONDS: 300
});

/*************************************************
 * ИНФРАСТРУКТУРА: ДОСТУП К ТАБЛИЦАМ
 *************************************************/
/** @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} */
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Получает лист по имени.
 * @param {string} name Имя листа.
 * @param {boolean} [createIfMissing=false] Создать лист, если его нет.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null}
 */
function getSheet(name, createIfMissing = false) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(name) || (createIfMissing ? ss.insertSheet(name) : null);
}

/*************************************************
 * ТОЧКА ВХОДА WEB-APP
 *************************************************/
/** Рендерит интерфейс веб-приложения. */
function doGet() {
  let output = HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('ДБВv5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  // setFaviconUrl не принимает data:image/...;base64. Нужен только публичный HTTPS URL.
  if (/^https:\/\//i.test(WEBAPP_FAVICON_URL)) {
    output = output.setFaviconUrl(WEBAPP_FAVICON_URL);
  }

  return output;
}

/**
 * Возвращает HTML-контент файла.
 * @param {string} name Имя HTML-файла.
 * @returns {string}
 */
function getHtmlFile(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/*************************************************
 * ЛОГИРОВАНИЕ
 *************************************************/
const LOG_COLUMNS = Object.freeze([
  'Дата/время входа',
  'Логин',
  'Пароль',
  'СНИЛС',
  'IP',
  'Устройство',
  'Браузер',
  'Статус входа'
]);
const LOG_MAX_AGE_MINUTES = 30;

/**
 * Подготавливает лист логов и гарантирует наличие актуальных заголовков.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null}
 */
function getLogSheet() {
  const sheet = getSheet(CONFIG.LOG_SHEET_NAME, true);
  if (!sheet) return null;

  if (sheet.getLastRow() === 0) {
    appendPlainLogRow(sheet, LOG_COLUMNS);
  } else {
    const currentHeader = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), LOG_COLUMNS.length)).getValues()[0];
    const needsHeaderUpdate = LOG_COLUMNS.some((title, index) => currentHeader[index] !== title);
    if (needsHeaderUpdate) {
      sheet.getRange(1, 1, 1, LOG_COLUMNS.length).setNumberFormat('@').setValues([LOG_COLUMNS]);
    }
  }

  return sheet;
}

/**
 * Возвращает дату и время в текстовом виде для многострочного лога.
 * @param {Date} value Дата.
 * @returns {string}
 */
function formatLogDateTime(value) {
  return Utilities.formatDate(value, CONFIG.TIMEZONE, `${CONFIG.DATE_FORMAT} HH:mm:ss`);
}

/**
 * Форматирует значение для конкретной колонки лога без добавления лишнего времени к паролю/дате рождения.
 * @param {number} col Индекс колонки лога.
 * @param {*} value Значение ячейки.
 * @returns {string}
 */
function formatLogCellValue(col, value) {
  if (value instanceof Date) {
    return col === 0 ? formatLogDateTime(value) : formatCellValue(value);
  }
  return String(value ?? '');
}

/**
 * Добавляет строку в лог как текст, чтобы Google Sheets не преобразовывал даты рождения в дату-время.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист логов.
 * @param {string[]} values Значения строки.
 */
function appendPlainLogRow(sheet, values) {
  const rowIndex = sheet.getLastRow() + 1;
  const range = sheet.getRange(rowIndex, 1, 1, LOG_COLUMNS.length);
  range.setNumberFormat('@').setValues([values]).setWrap(true);
}

/**
 * Разбирает дату первой строки лога в формате dd.MM.yyyy HH:mm:ss.
 * @param {*} value Значение ячейки с датой.
 * @returns {Date | null}
 */
function parseLogDateTime(value) {
  if (value instanceof Date) return value;

  const firstLine = splitLogLines(value)[0] || '';
  const match = firstLine.match(/^(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
  if (!match) return null;

  const [, day, month, year, hours, minutes, seconds] = match;
  return new Date(Number(year), Number(month) - 1, Number(day), Number(hours), Number(minutes), Number(seconds));
}

/**
 * Проверяет, можно ли использовать IP для поиска строки лога.
 * @param {*} value IP из clientInfo.
 * @returns {boolean}
 */
function isRecognizedIp(value) {
  const ip = String(value || '').trim().toLowerCase();
  return Boolean(ip) && ip !== 'не определён' && ip !== 'не определен';
}

/**
 * Делит строковое значение ячейки на логические строки.
 * @param {*} value Значение ячейки.
 * @returns {string[]}
 */
function splitLogLines(value) {
  const text = String(value ?? '');
  return text === '' ? [] : text.split('\n');
}

/**
 * Берет последнюю непустую строку из ячейки лога.
 * @param {*} value Значение ячейки.
 * @returns {string}
 */
function getLastLogLine(value) {
  const lines = splitLogLines(value).filter(line => line !== '');
  return lines.length ? lines[lines.length - 1] : '';
}

/**
 * Делает все предыдущие строки зачеркнутыми, а последнюю строку оставляет обычной.
 * @param {string[]} lines Строки ячейки.
 * @returns {GoogleAppsScript.Spreadsheet.RichTextValue}
 */
function buildLogRichText(lines) {
  const text = lines.join('\n');
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const normalStyle = SpreadsheetApp.newTextStyle().setStrikethrough(false).build();
  const strikeStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();

  if (text.length > 0) {
    builder.setTextStyle(0, text.length, normalStyle);
  }

  if (lines.length > 1) {
    const previousLength = lines.slice(0, -1).join('\n').length;
    if (previousLength > 0) {
      builder.setTextStyle(0, previousLength, strikeStyle);
    }
  }

  return builder.build();
}

/**
 * Ищет строку открытия сайта: сначала по IP, затем по ФИО, только за последние 30 минут.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист логов.
 * @param {{login?:string,clientInfo?:Object}} payload Данные для поиска.
 * @returns {number} Номер строки или -1.
 */
function findRecentLogRow(sheet, { login = '', clientInfo = {} }) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 1, lastRow - 1, LOG_COLUMNS.length).getValues();
  const cutoff = Date.now() - LOG_MAX_AGE_MINUTES * 60 * 1000;
  const targetIp = String(clientInfo.ip || '').trim();
  const canUseIp = isRecognizedIp(targetIp);
  const targetLogin = normalizeLogin(login);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const dateValue = parseLogDateTime(row[0]);
    if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime()) || dateValue.getTime() < cutoff) {
      continue;
    }

    if (canUseIp && getLastLogLine(row[4]) === targetIp) {
      return i + 2;
    }
  }

  if (canUseIp || !targetLogin) return -1;
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const dateValue = parseLogDateTime(row[0]);
    if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime()) || dateValue.getTime() < cutoff) {
      continue;
    }

    if (normalizeLogin(getLastLogLine(row[1])) === targetLogin) {
      return i + 2;
    }
  }

  return -1;
}

/**
 * Добавляет новую логическую строку к найденной записи, зачеркивая предыдущие данные.
 * Если в одной ячейке появляется новая строка, перенос добавляется во все остальные ячейки строки.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист логов.
 * @param {number} rowIndex Номер строки.
 * @param {string[]} newValues Новые значения по колонкам лога.
 */
function appendLogLine(sheet, rowIndex, newValues) {
  const range = sheet.getRange(rowIndex, 1, 1, LOG_COLUMNS.length);
  const oldValues = range.getValues()[0];
  const richValues = [];

  for (let col = 0; col < LOG_COLUMNS.length; col++) {
    const oldText = formatLogCellValue(col, oldValues[col]);
    const oldLines = oldText === '' ? [''] : oldText.split('\n');
    const nextValue = newValues[col] == null ? '' : String(newValues[col]);
    const lines = oldLines.length ? oldLines.concat([nextValue]) : [nextValue];
    richValues.push(buildLogRichText(lines));
  }

  range.setRichTextValues([richValues]).setWrap(true);
}

/**
 * Создает первичную строку при открытии сайта.
 * @param {{login?:string,password?:string,snils?:string,clientInfo?:Object}} payload Данные открытия.
 */
function logPageOpen({ login = '', password = '', snils = '', clientInfo = {} } = {}) {
  const sheet = getLogSheet();
  if (!sheet) return;

  appendPlainLogRow(sheet, [
    formatLogDateTime(new Date()),
    login,
    password,
    snils,
    clientInfo.ip || '',
    clientInfo.device || '',
    clientInfo.browser || '',
    ''
  ]);
}

/**
 * Пишет результат авторизации в строку открытия сайта, найденную по IP или ФИО за последние 30 минут.
 * @param {{login?:string,password?:string,snils?:string,clientInfo?:Object,status:string}} payload
 */
function logAccess({ login = '', password = '', snils = '', clientInfo = {}, status }) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getLogSheet();
    if (!sheet) return;

    const values = [
    formatLogDateTime(new Date()),
    login,
    password,
    snils,
    clientInfo.ip || '',
    clientInfo.device || '',
    clientInfo.browser || '',
    status || ''
  ];
  const rowIndex = findRecentLogRow(sheet, { login, clientInfo });

  if (rowIndex > 0) {
      appendLogLine(sheet, rowIndex, values);
      return;
    }

    appendPlainLogRow(sheet, values);
  } finally {
    lock.releaseLock();
  }
}


/**
 * Логирует попытку входа, пропуская технические фоновые проверки.
 * @param {{login?:string,password?:string,snils?:string,clientInfo?:Object,status:string}} payload
 */
function logAuthAttempt(payload) {
  if (payload.clientInfo && payload.clientInfo.silent) return;
  logAccess(payload);
}

/**
 * Логирует клик по кнопке открытия формы.
 * @param {string} login Логин пользователя.
 */
function logEditFormClick(login) {
  logAccess({ login, status: 'Нажал: Заполнить / редактировать' });
}

/**
 * Логирует нажатие кнопки входа.
 * @param {string} login Логин пользователя.
 */
function logLoginButtonClick(login) {
  logAccess({ login, status: 'Нажал: Войти' });
}

/**
 * Логирует переход в раздел.
 * @param {string} login Логин пользователя.
 * @param {string} section Ключ раздела.
 */
function logSectionVisit(login, section) {
  const sectionMap = {
    docs: 'ДБВv5 Документы',
    attendance: 'ДБВv5 Посещаемость',
    gear: 'ДБВv5 Снаряжение'
  };
  const sectionName = sectionMap[section] || section;
  logAccess({ login, status: `Перешёл в раздел ${sectionName}` });
}

/**
 * Логирует клик по кнопке заполнения формы.
 * @param {string} login Логин пользователя.
 */
function logFillFormClick(login) {
  logAccess({ login, status: 'Нажал: Заполнить форму' });
}

/**
 * Логирует открытие формы в отдельной вкладке.
 * @param {string} login Логин пользователя.
 */
function logExternalFormClick(login) {
  logAccess({ login, status: 'Нажал: Открыть форму в отдельной вкладке' });
}

/*************************************************
 * ДОКУМЕНТЫ: АВТОРИЗАЦИЯ И ПОДГОТОВКА ДАННЫХ
 *************************************************/
/**
 * Находит индексы колонок, доступных к показу (желтые заголовки).
 * @param {string[]} headerColors Цвета заголовков.
 * @returns {number[]}
 */
function getAllowedColumnIndexes(headerColors) {
  const result = [];
  for (let i = 0; i < headerColors.length; i++) {
    if (headerColors[i] === CONFIG.YELLOW) result.push(i);
  }
  return result;
}

/**
 * Возвращает индексы колонок логина и пароля.
 * @param {string[]} header Заголовки таблицы.
 * @returns {{loginCol:number, passCol:number}}
 */
function getAuthColumnIndexes(header) {
  const loginCol = header.indexOf('Фамилия Имя Отчество (С)');
  const passCol = header.indexOf('Дата рождения (С)');

  if (loginCol === -1 || passCol === -1) {
    throw new Error('AUTH_COLUMNS_NOT_FOUND');
  }

  return { loginCol, passCol };
}


/**
 * Возвращает индекс колонки СНИЛС.
 * @param {string[]} header Заголовки таблицы.
 * @returns {number}
 */
function getSnilsColumnIndex(header) {
  return header.indexOf('Снилс: номер');
}

/**
 * Пакетно загружает колонки авторизации из минимального диапазона.
 * Это снижает число обращений к Spreadsheet API.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowCount
 * @param {{loginCol:number,passCol:number}} authCols
 * @param {number} snilsCol
 * @returns {{logins:Array<Array<*>>,passwords:Array<Array<*>>,snilsValues:Array<Array<*>>|null}}
 */
function loadAuthColumns(sheet, rowCount, authCols, snilsCol) {
  const requestedCols = [authCols.loginCol, authCols.passCol];
  if (snilsCol >= 0) requestedCols.push(snilsCol);

  const minCol = Math.min(...requestedCols);
  const maxCol = Math.max(...requestedCols);
  const width = maxCol - minCol + 1;
  const block = sheet.getRange(2, minCol + 1, rowCount, width).getValues();

  const rel = col => col - minCol;
  const logins = block.map(row => [row[rel(authCols.loginCol)]]);
  const passwords = block.map(row => [row[rel(authCols.passCol)]]);
  const snilsValues = snilsCol >= 0 ? block.map(row => [row[rel(snilsCol)]]) : null;

  return { logins, passwords, snilsValues };
}

/**
 * Нормализует СНИЛС к формату 000-000-000-00.
 * @param {*} value Исходный СНИЛС.
 * @returns {string}
 */
function normalizeSnils(value) {
  const digits = String(value || '').replace(/\D/g, '').slice(0, 11);
  if (digits.length !== 11) return '';
  return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 9)}-${digits.slice(9, 11)}`;
}

/**
 * Нормализует логин/ФИО для корректного сравнения.
 * @param {*} value Исходное значение.
 * @returns {string}
 */
function normalizeLogin(value) {
  return String(value || '').trim().toLowerCase().replace(/ё/g, 'е');
}

/**
 * Приводит значение ячейки к строке (включая дату).
 * @param {*} value Значение ячейки.
 * @returns {string}
 */
function formatCellValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  }
  return String(value ?? '');
}

/**
 * Собирает данные строки только по разрешенным колонкам.
 * @param {Array<*>} row Значения строки.
 * @param {string[]} header Заголовки.
 * @param {string[]} backgrounds Цвета ячеек строки.
 * @param {number[]} allowedCols Индексы разрешенных колонок.
 */
function prepareRowForClient(row, header, backgrounds, allowedCols) {
  return {
    header: allowedCols.map(i => header[i]),
    row: allowedCols.map(i => formatCellValue(row[i])),
    colors: allowedCols.map(i => backgrounds[i])
  };
}

/**
 * Техническое логирование этапов выполнения (для поиска зависаний).
 * @param {string} stage
 * @param {number} startedAt
 */
function logStage(stage, startedAt) {
  const elapsed = Date.now() - startedAt;
  Logger.log(`[AUTH][${elapsed}ms] ${stage}`);
}

/**
 * Проверяет логин/дату рождения и возвращает персональные данные для UI.
 * @param {string} login ФИО.
 * @param {string} password Дата рождения в формате ДД.ММ.ГГГГ.
 * @param {Object} [clientInfo={}] Данные об устройстве.
 * @returns {Object}
 */
function checkLogin(login, password, clientInfo = {}, snils = '') {
  const startedAt = Date.now();
  logStage('Начало checkLogin', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);

  if (!sheet) {
    logAuthAttempt({ login, password, snils, clientInfo, status: 'Лист не найден' });
    return { error: 'Лист с результатами не найден.' };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  logStage(`Получены границы листа: rows=${lastRow}, cols=${lastCol}`, startedAt);

  if (lastRow < 2 || lastCol < 1) {
    return { error: 'Таблица пуста.' };
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const headerColors = sheet.getRange(1, 1, 1, lastCol).getBackgrounds()[0];
  logStage('Загружены заголовки и цвета заголовков', startedAt);

  let authCols;
  let allowedCols;

  try {
    authCols = getAuthColumnIndexes(header);
    allowedCols = getAllowedColumnIndexes(headerColors);
  } catch (_error) {
    logAuthAttempt({ login, password, snils, clientInfo, status: 'Ошибка конфигурации столбцов' });
    return { error: 'Ошибка структуры таблицы.' };
  }

  const rowCount = lastRow - 1;
  const normalizedLogin = normalizeLogin(login);
  const snilsCol = getSnilsColumnIndex(header);
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, rowCount, authCols, snilsCol);
  const expectedSnils = normalizeSnils(snils);
  logStage('Загружены только колонки авторизации/СНИЛС', startedAt);

  for (let i = 0; i < rowCount; i++) {
    const rowLogin = normalizeLogin(logins[i][0]);
    const rowPassword = formatCellValue(passwords[i][0]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = snilsValues ? normalizeSnils(snilsValues[i][0]) : '';
      if (rowSnils && rowSnils !== expectedSnils) {
        const snilsVisible = Boolean(clientInfo.snilsVisible);
        const snilsStatus = expectedSnils && snilsVisible ? 'Неверный СНИЛС' : 'Требуется ввод СНИЛС';
        logAuthAttempt({ login, password, snils, clientInfo, status: snilsStatus });
        logStage(expectedSnils && snilsVisible ? 'Совпадение найдено, СНИЛС неверный' : 'Совпадение найдено, требуется СНИЛС', startedAt);
        return { requiresSnils: true, snilsError: expectedSnils && snilsVisible ? 'invalid' : 'required' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      logAuthAttempt({ login, password, snils, clientInfo, status: 'Удачный вход' });
      logStage('Совпадение найдено, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  logAuthAttempt({ login, password, snils, clientInfo, status: 'Неудачный вход: ФИО/дата' });
  logStage('Совпадение не найдено', startedAt);
  return { error: 'Неправильно введены ФИО или дата рождения.' };
}


/**
 * Второй фактор: проверка СНИЛС после успешного совпадения ФИО и даты рождения.
 * @param {string} login ФИО.
 * @param {string} password Дата рождения.
 * @param {string} snils СНИЛС из формы.
 * @param {Object} [clientInfo={}] Данные клиента.
 * @returns {Object}
 */
function verifySnils(login, password, snils, clientInfo = {}) {
  const startedAt = Date.now();
  logStage('Начало verifySnils', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) {
    return { error: 'Лист с результатами не найден.' };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    return { error: 'Таблица пуста.' };
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const headerColors = sheet.getRange(1, 1, 1, lastCol).getBackgrounds()[0];

  let authCols;
  let allowedCols;
  try {
    authCols = getAuthColumnIndexes(header);
    allowedCols = getAllowedColumnIndexes(headerColors);
  } catch (_error) {
    return { error: 'Ошибка структуры таблицы.' };
  }

  const snilsCol = getSnilsColumnIndex(header);
  if (snilsCol === -1) {
    return { error: 'Ошибка структуры таблицы.' };
  }

  const rowCount = lastRow - 1;
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, rowCount, authCols, snilsCol);
  logStage('Загружены колонки для проверки СНИЛС', startedAt);

  const normalizedLogin = normalizeLogin(login);
  const expectedSnils = normalizeSnils(snils);

  for (let i = 0; i < rowCount; i++) {
    const rowLogin = normalizeLogin(logins[i][0]);
    const rowPassword = formatCellValue(passwords[i][0]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = normalizeSnils(snilsValues[i][0]);
      if (!rowSnils) {
        const rowIndex = i + 2;
        const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
        const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
        logAuthAttempt({ login, password, snils, clientInfo, status: 'Удачный вход без СНИЛС' });
        return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
      }

      if (rowSnils !== expectedSnils) {
        logAuthAttempt({ login, password, snils, clientInfo, status: 'Неверный СНИЛС' });
        return { error: 'Неверный СНИЛС.' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      logAuthAttempt({ login, password, snils, clientInfo, status: 'Удачный вход по СНИЛС' });
      logStage('СНИЛС подтвержден, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  return { error: 'Неправильно введены ФИО или дата рождения.' };
}

/*************************************************
 * ПОСЕЩАЕМОСТЬ: ПОДБОР ФИО
 *************************************************/
/**
 * Подбирает полные ФИО по списку сокращенных записей.
 * @param {string[]} inputs Введенные сокращенные ФИО.
 * @returns {Array<Object|null>}
 */
function findNames(inputs) {
  const fullNames = loadFullNames();
  return inputs.map(input => processInput(input, fullNames));
}

/**
 * Загружает полные ФИО из первого столбца листа результатов.
 * Для ускорения используется краткоживущий кэш ScriptCache.
 * @returns {Array<{original:string,last:string,first:string,middle:string}>}
 */
function loadFullNames() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.NAMES_CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) return [];

  const values = sheet
    .getRange(1, 1, sheet.getLastRow(), 1)
    .getValues()
    .flat()
    .filter(String)
    .map(normalizeFullName);

  cache.put(CONFIG.NAMES_CACHE_KEY, JSON.stringify(values), CONFIG.NAMES_CACHE_TTL_SECONDS);
  return values;
}

/**
 * Подбирает лучшие варианты полного ФИО для одного ввода.
 * @param {string} input Сокращенное ФИО.
 * @param {Array<Object>} fullNames Полный справочник ФИО.
 * @returns {{selected:string,options:string[]}|null}
 */
function processInput(input, fullNames) {
  const short = normalizeShortName(input);
  if (!short) return null;

  const maxErrors = 2;
  const matches = fullNames
    .map(full => calculateMatch(short, full, maxErrors))
    .filter(Boolean)
    .sort((a, b) => a.totalCost - b.totalCost);

  if (!matches.length) return null;

  const exactLast = matches.filter(m => m.lastCost === 0);
  const selected = exactLast.length === 1 ? exactLast[0].original : matches[0].original;

  const top = matches.slice(0, 3);
  if (!top.some(m => m.original === selected)) {
    const selectedMatch = matches.find(m => m.original === selected);
    if (selectedMatch) {
      top.pop();
      top.unshift(selectedMatch);
    }
  }

  const bestCost = top.length ? top[0].totalCost : null;
  const sameBestCount = bestCost === null ? 0 : top.filter(m => m.totalCost === bestCost).length;
  const shouldAutoSelect = sameBestCount <= 1;

  return {
    selected: shouldAutoSelect ? selected : '',
    options: top.map(m => m.original)
  };
}


/**
 * Возвращает список последних тренировок из внешней таблицы ответов.
 * Порядок: в таблице сверху-вниз, на сайте снизу-вверх.
 * @returns {Array<{timestamp:string,date:string,coach:string,place:string,fio:string,fioGroups:Array<{group:string,names:string[]}>}>}
 */
function getTrainingHistory() {
  const ss = SpreadsheetApp.openById('1K1TtjIL2retzFoXBlQaePKbeKIEkMZZedZX-Ans4VjY');
  const sheet = ss.getSheetByName('ОтветыV5');
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(String);
  const tsCol = header.indexOf('Отметка времени');
  const dateCol = header.indexOf('Дата тренировки');
  const coachCol = header.indexOf('Тренер присутствовал:');
  const placeCol = header.indexOf('Место проведения занятия');

  if (tsCol === -1 || dateCol === -1 || coachCol === -1 || placeCol === -1) return [];

  const fioGroupMap = loadTrainingGroupMap();
  const startCol = 12; // M
  const endCol = 32;   // AG

  const rows = values.slice(1)
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      const fioNames = extractFioNamesFromRow(row, startCol, endCol);
      const fioGroups = buildFioGroups(fioNames, fioGroupMap);
      const fioMerged = fioNames.join('\n');

      return {
        timestamp: formatTrainingCell(row[tsCol]),
        date: formatTrainingCell(row[dateCol]),
        coach: formatTrainingCell(row[coachCol]),
        place: formatTrainingCell(row[placeCol]),
        fio: fioMerged,
        fioGroups
      };
    })
    .reverse();

  return rows;
}

/**
 * Загружает соответствие ФИО -> тренировочная группа.
 * @returns {Object<string,string>}
 */
function loadTrainingGroupMap() {
  const ss = SpreadsheetApp.openById('1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ');
  const sheet = ss.getSheetByName('Результат');
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0].map(String);
  const fioCol = header.indexOf('Фамилия Имя Отчество (С)');
  const groupCol = header.indexOf('Тренировочная группа');
  if (fioCol === -1 || groupCol === -1) return {};

  const map = {};
  values.slice(1).forEach(row => {
    const fio = String(row[fioCol] || '').trim();
    const group = String(row[groupCol] || '').trim();
    if (!fio) return;
    map[normalizeTrainingName(fio)] = group || 'Без группы';
  });

  return map;
}

/**
 * Извлекает ФИО из диапазона столбцов строки ответа.
 * @param {Array<*>} row
 * @param {number} startCol
 * @param {number} endCol
 * @returns {string[]}
 */
function extractFioNamesFromRow(row, startCol, endCol) {
  const names = [];

  row.slice(startCol, endCol + 1).forEach(value => {
    const parts = String(value || '')
      .replace(/\r/g, '\n')
      .replace(/,\s*/g, '\n')
      .split(/\n+/)
      .map(s => s.trim())
      .filter(Boolean);

    parts.forEach(name => names.push(name));
  });

  return names;
}

/**
 * Группирует ФИО по тренировочным группам.
 * @param {string[]} fioNames
 * @param {Object<string,string>} fioGroupMap
 * @returns {Array<{group:string,names:string[]}>}
 */
function buildFioGroups(fioNames, fioGroupMap) {
  const groups = {};
  const order = [];

  fioNames.forEach(name => {
    const normalized = normalizeTrainingName(name);
    const group = fioGroupMap[normalized] || 'Без группы';
    if (!groups[group]) {
      groups[group] = [];
      order.push(group);
    }
    groups[group].push(name);
  });

  return order.map(group => ({ group, names: groups[group] }));
}

/**
 * Нормализация ФИО для сопоставления тренировочной группы.
 * @param {string} value
 * @returns {string}
 */
function normalizeTrainingName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Форматирует значение ячейки для блока истории тренировок.
 * @param {*} value
 * @returns {string}
 */
function formatTrainingCell(value) {
  if (value instanceof Date) {
    const hasTime = value.getHours() !== 0 || value.getMinutes() !== 0 || value.getSeconds() !== 0;
    const pattern = hasTime ? 'dd.MM.yyyy HH:mm:ss' : CONFIG.DATE_FORMAT;
    return Utilities.formatDate(value, CONFIG.TIMEZONE, pattern);
  }
  return String(value ?? '');
}

/**
 * Базовая нормализация текстового ввода.
 * @param {*} text Исходный текст.
 * @returns {string}
 */
function normalize(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\./g, '')
    .replace(/\s+/g, '');
}

/**
 * Делит полное ФИО на части: фамилия, имя, отчество.
 * @param {string} text Полное ФИО.
 */
function normalizeFullName(text) {
  const clean = normalize(text);
  const m = clean.match(/^([а-я]+)([а-я]+)?([а-я]+)?$/) || [];

  return {
    original: text,
    last: m[1] || '',
    first: m[2] || '',
    middle: m[3] || ''
  };
}

/**
 * Нормализует сокращенный ввод вида "ФамилияИ".
 * @param {string} text Ввод пользователя.
 * @returns {{last:string, tail:string}|null}
 */
function normalizeShortName(text) {
  if (!text) return null;

  const clean = normalize(text);
  const m = clean.match(/^([а-я]+)([а-я]*)$/);
  if (!m) return null;

  return { last: m[1], tail: m[2] };
}

/**
 * Стоимость fuzzy-сопоставления префикса (Левенштейн) в пределах maxErrors.
 * @param {string} text Эталонный текст.
 * @param {string} pattern Шаблон.
 * @param {number} maxErrors Допустимый бюджет ошибок.
 * @returns {number}
 */
function fuzzyPrefixCost(text, pattern, maxErrors) {
  if (!pattern) return 0;
  if (!text) return Infinity;

  let min = Infinity;
  const minLen = Math.max(1, pattern.length - maxErrors);
  const maxLen = Math.min(text.length, pattern.length + maxErrors);

  for (let len = minLen; len <= maxLen; len++) {
    const d = levenshtein(text.slice(0, len), pattern);
    if (d < min) min = d;
  }

  return min;
}

/**
 * Расстояние Левенштейна между строками.
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
  }

  return dp[m][n];
}

/**
 * Считает качество совпадения сокращенного ФИО с полным.
 * @param {{last:string,tail:string}} short Сокращенный ввод.
 * @param {{original:string,last:string,first:string,middle:string}} full Полное ФИО.
 * @param {number} maxErrors Бюджет ошибок.
 * @returns {{original:string,lastCost:number,totalCost:number}|null}
 */
function calculateMatch(short, full, maxErrors) {
  let budget = maxErrors;

  const lastCost = fuzzyPrefixCost(full.last, short.last, budget);
  if (lastCost > budget) return null;
  budget -= lastCost;

  let tailCost = 0;
  if (short.tail) {
    const costs = [
      fuzzyPrefixCost(full.first, short.tail, budget),
      fuzzyPrefixCost(full.middle, short.tail, budget),
      fuzzyPrefixCost(full.first + full.middle, short.tail, budget)
    ];
    tailCost = Math.min(...costs);
    if (tailCost > budget) return null;
  }

  return {
    original: full.original,
    lastCost,
    totalCost: lastCost + tailCost
  };
}

/**
 * Обновляет одну ячейку на листе "Результат" по названию колонки для авторизованной строки.
 * @param {string} login
 * @param {string} password
 * @param {string} snils
 * @param {string} columnName
 * @param {string} value
 * @returns {{ok:boolean}}
 */
function updateResultCell(login, password, snils, columnName, value) {
  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) throw new Error('Лист с результатами не найден.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('Таблица пуста.');

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const targetCol = header.indexOf(String(columnName || ''));
  if (targetCol === -1) throw new Error('Колонка не найдена.');

  const authCols = getAuthColumnIndexes(header);
  const snilsCol = getSnilsColumnIndex(header);
  const rowCount = lastRow - 1;
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, rowCount, authCols, snilsCol);

  const normalizedLogin = normalizeLogin(login);
  const normalizedSnils = normalizeSnils(snils);
  const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm:ss');
  const newValue = String(value ?? '');

  for (let i = 0; i < rowCount; i++) {
    const rowLogin = normalizeLogin(logins[i][0]);
    const rowPassword = formatCellValue(passwords[i][0]).trim();
    if (rowLogin !== normalizedLogin || rowPassword !== String(password || '').trim()) continue;

    const rowSnils = normalizeSnils(snilsValues[i][0]);
    if (rowSnils && normalizedSnils && rowSnils !== normalizedSnils) continue;

    const rowIndex = i + 2;
    const cell = sheet.getRange(rowIndex, targetCol + 1);
    cell.setValue(newValue);

    if (String(columnName || '').trim() === 'Школа') {
      const schoolUpdatedCol = header.indexOf('Дата обн. инф. о школе (С)');
      if (schoolUpdatedCol !== -1) {
        sheet.getRange(rowIndex, schoolUpdatedCol + 1).setValue(now);
      }
    }

    const historyLine = `S: ${now}, ${newValue}`;
    const existingNote = String(cell.getNote() || '');
    const nextNote = existingNote ? `${historyLine}\n${existingNote}` : historyLine;

    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        cell.setNote(nextNote);
        break;
      } catch (e) {
        if (attempt === 2) throw e;
        Utilities.sleep(150);
      }
    }

    return { ok: true };
  }

  throw new Error('Строка для обновления не найдена.');
}
