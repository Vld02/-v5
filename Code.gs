/*************************************************
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 *************************************************/
const WEBAPP_FAVICON_URL = 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.ico'; // Вставьте прямую HTTPS-ссылку на PNG/ICO для вкладки в обёртке Google Script.

const SERVER_CONFIG = Object.freeze({
  drive: Object.freeze({ usersRootFolderId: '1AyjWNspWbBVswPdrSy0M-JEbvZBsjq1' }),
  userFolderTemplate: '{Фамилия} {Имя} {Отчество} {Дата рождения (разряд)}'
});

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
 * Логирует клик по кнопке заполнения формы тренировки.
 * @param {string} login Логин пользователя.
 */
function logFillFormClick(login) {
  logAccess({ login, status: 'Нажал: Заполнить форму' });
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


/* ============================================================
   НАБОР ПРАВИЛ РЕДАКТИРОВАНИЯ — СЕРВЕРНЫЙ ИСТОЧНИК ПРОВЕРКИ

   Серверная копия той же архитектуры, которую использует браузер
   для удобства ввода. В Apps Script HTML и Code.gs выполняются в
   разных контекстах, поэтому сервер обязан иметь собственный доступ
   к конфигурации и никогда не доверять объектам, изменённым в браузере.

   ВАЖНО: именно этот блок защищает запись в Google Таблицу. Перед
   setValue() updateResultCell() проверяет существование поля,
   editable, required и regex по EDIT_CONFIG.fields + EDIT_CONFIG.rules.
   ============================================================ */
const EDIT_CONFIG = Object.freeze({
  rules: Object.freeze({
    TEXT: { title: 'Текст', placeholder: '', regex: '^.*$', special: '' },
    SUGGEST_TEXT: { title: 'Текст из подсказок', placeholder: '', regex: '^.*$', special: 'suggest' },
    FULL_NAME_RU: { title: 'ФИО', placeholder: 'Иванов Иван Иванович', regex: '^\\s*[А-ЯЁ][а-яё]+\\s+[А-ЯЁ][а-яё]+\\s+[А-ЯЁ][а-яё]+\\s*$', special: 'fullname' },
    YEAR: { title: 'Год', placeholder: '2024', regex: '^\\d{4}$', special: 'year' },
    CLASS_COURSE: { title: 'Класс / курс', placeholder: '7', regex: '^(?:[0-9]|1[01]|I|II|III|IV|V|VI)$', special: 'classCourse' },
    RU_UPPER_LETTER: { title: 'Русская заглавная буква', placeholder: 'А', regex: '^[А-ЯЁ]$', special: 'singleRuUpper' },
    PHONE_RU: { title: 'Номер телефона', placeholder: '+7 999 123-45-67', regex: '^\\+7\\s\\d{3}\\s\\d{3}-\\d{2}-\\d{2}$', special: 'phoneRu' },
    EMAIL: { title: 'Электронная почта', placeholder: 'name@example.ru', regex: '^[A-Za-z0-9.!#$%&\'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\\.[A-Za-z0-9-]+)+$', special: 'email' },
    CERTIFICATE_RU: { title: 'Свидетельство о рождении', placeholder: 'IV-АБ № 123456', regex: '^([VIX]{1,4}-[А-ЯЁ]{1,3}\\s*№\\s*\\d{6})$', special: 'certificate' },
    DATE_RU: { title: 'Дата', placeholder: 'ДД.ММ.ГГГГ', regex: '^\\d{2}\\.\\d{2}\\.\\d{4}$', special: 'date' },
    SNILS: { title: 'СНИЛС', placeholder: '000-000-000-00', regex: '^\\d{3}[-\\s]?\\d{3}[-\\s]?\\d{3}[-\\s]?\\d{2}$', special: 'snils' },
    PASSPORT_RU: { title: 'Паспорт РФ', placeholder: '12 34 567890', regex: '^\\d{2}\\s\\d{2}\\s\\d{6}$', special: 'passportRu' },
    PASSPORT_DIVISION_CODE: { title: 'Код подразделения', placeholder: '123-456', regex: '^\\d{3}-\\d{3}$', special: 'passportDivisionCode' },
    MED_POLICY_NUMBER: { title: 'Номер медполиса', placeholder: '1234 5678 9012 3456', regex: '^\\d{4}\\s\\d{4}\\s\\d{4}\\s\\d{4}$', special: 'medPolicyNumber' },
    MGFSO_ID: { title: 'ID МГФСО', placeholder: '1234567', regex: '^\\d{7}$', special: 'mgfsoId' }
  }),

  /* ============================================================
   СООТВЕТСТВИЕ ПОЛЕЙ И ПРАВИЛ — СЕРВЕРНАЯ КАРТА ДОСТУПА

   Этот список определяет, какие реальные заголовки таблицы можно
   менять. Если заголовка здесь нет или editable: false, сервер
   откажет даже при ручном вызове updateResultCell() из DevTools.
   ============================================================ */
  fields: Object.freeze({
    'Фамилия Имя Отчество (С)': { editable: true, rule: 'FULL_NAME_RU', required: true, isIdentityField: true, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Дата рождения (С)': { editable: true, rule: 'DATE_RU', required: true, isIdentityField: true, description: 'Дата рождения в формате ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Месяц рождения (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Год набора': { editable: true, rule: 'YEAR', required: false, isIdentityField: false, description: 'Год в динамическом диапазоне 1950 — текущий год + 1.', example: '2024' },
    'Пол (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Школа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, onSave: 'UPDATE_SCHOOL_DATE', description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Школы', startRow: 2 } },
    'Класс / курс': { editable: true, rule: 'CLASS_COURSE', required: false, isIdentityField: false, description: '0-11 или I-VI.', example: '7' },
    'Литера класса (буква)': { editable: true, rule: 'RU_UPPER_LETTER', required: false, isIdentityField: false, description: 'Одна заглавная русская буква.', example: 'А' },
    'Директор школы: Фамилия Имя Отчество': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Адрес регистрации / Прописка (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Телефон +7 (С)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (С)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Фамилия Имя Отчество (П)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (П)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (П)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (П)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (П)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Кем выдан (П)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (П)', startRow: 2 } },
    'Паспорт: Когда выдан (П)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (П)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Фамилия Имя Отчество (М)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (М)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (М)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (М)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (М)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (М)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (М)', startRow: 2 } },
    'Паспорт: Когда выдан (М)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (М)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Фамилия Имя Отчество (Д)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (Д)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (Д)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (Д)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (Д)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (Д)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (Д)', startRow: 2 } },
    'Паспорт: Когда выдан (Д)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (Д)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Свидетельство: Серия, номер (С)': { editable: true, rule: 'CERTIFICATE_RU', required: false, isIdentityField: false, description: 'Серия и номер свидетельства.', example: 'IV-АБ № 123456' },
    'Паспорт: Серия, номер (С)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Свидетельство: Кем выдан (С)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Свидетельство: Кем выдан (С)', startRow: 2 } },
    'Паспорт: Кем выдан': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан', startRow: 2 } },
    'Паспорт или Свидетельство: Кем выдан (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Свидетельство: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт или Свидетельство: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Код подразделения (С)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Снилс: номер': { editable: true, rule: 'SNILS', required: false, isIdentityField: true, description: 'СНИЛС из 11 цифр.', example: '000-000-000-00' },
    'Полис: Страховая компания': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Полис: Страховая компания', startRow: 2 } },
    'Тренировочная группа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Группы тренировки', startRow: 5 } },
    'МГФСО группа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Группы МГФСО', startRow: 2 } },
    'Тренер МГФСО': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Тренер МГФСО', startRow: 2 } },
    'Разряд': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Список разрядов', startRow: 2 } },
    'Мед полис: номер': { editable: true, rule: 'MED_POLICY_NUMBER', required: false, isIdentityField: false, description: 'Номер медицинского полиса в формате 1234 5678 9012 3456.', example: '1234 5678 9012 3456' },
    'Дата зачисления в МГФСО': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Укажите дату зачисления в МГФСО.', example: '15.08.2024' },
    'Дата получения разряда': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Страховка: Действительна до': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Страховка: Номер и компаия': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'ID номер МГФСО': { editable: true, rule: 'MGFSO_ID', required: false, isIdentityField: false, description: 'ID МГФСО из 7 цифр.', example: '1234567' },
    'Мед допуск до': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'РУСАДА': { editable: true, rule: 'YEAR', required: false, isIdentityField: false, description: 'Год в динамическом диапазоне 1950 — текущий год + 1.', example: '2024' },
    'Свидетельство: скан (C)': { editable: false, attachment: 'CERTIFICATE_C' },
    'Паспорт: скан (С)': { editable: false, attachment: 'PASSPORT_C' },
    'Снилс: Скан (C)': { editable: false, attachment: 'SNILS_C' },
    'Полис: Скан (C)': { editable: false, attachment: 'POLICY_C' },
    'Мед допуск: Скан': { editable: false, attachment: 'MED_CLEARANCE' },
    'Русада: Скан': { editable: false, attachment: 'RUSADA' },
    'Паспорт: Скан (П)': { editable: false, attachment: 'PASSPORT_P' },
    'Паспорт: Скан (М)': { editable: false, attachment: 'PASSPORT_M' },
    'Паспорт: Скан (Д)': { editable: false, attachment: 'PASSPORT_D' }
  }),
  attachments: Object.freeze({
    CERTIFICATE_C: { title: 'Свидетельство: скан (C)', enabled: true, linkColumn: 'Свидетельство: скан (C)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Свидетельство', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    PASSPORT_C: { title: 'Паспорт: скан (С)', enabled: true, linkColumn: 'Паспорт: скан (С)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Паспорт', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    SNILS_C: { title: 'Снилс: Скан (C)', enabled: true, linkColumn: 'Снилс: Скан (C)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - СНИЛС', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    POLICY_C: { title: 'Полис: Скан (C)', enabled: true, linkColumn: 'Полис: Скан (C)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Полис', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    MED_CLEARANCE: { title: 'Мед допуск: Скан', enabled: true, linkColumn: 'Мед допуск: Скан', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Мед допуск', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    RUSADA: { title: 'Русада: Скан', enabled: true, linkColumn: 'Русада: Скан', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - РУСАДА', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    PASSPORT_P: { title: 'Паспорт: Скан (П)', enabled: true, linkColumn: 'Паспорт: Скан (П)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Паспорт законного представителя', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    PASSPORT_M: { title: 'Паспорт: Скан (М)', enabled: true, linkColumn: 'Паспорт: Скан (М)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Паспорт законного представителя', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' },
    PASSPORT_D: { title: 'Паспорт: Скан (Д)', enabled: true, linkColumn: 'Паспорт: Скан (Д)', fileNameTemplate: '{Фамилия} {Имя} {Отчество} - Паспорт законного представителя', allowedExtensions: ['pdf', 'jpg', 'jpeg', 'png'], maxSizeMb: 20, replaceMode: 'version' }
  })
});

/**
 * Возвращает браузеру конфигурацию редактирования из единственного
 * серверного источника истины. Клиент использует её только для UI/UX;
 * updateResultCell() всё равно проверяет права напрямую по EDIT_CONFIG.
 * @returns {{rules:Object, fields:Object}}
 */
function getEditConfig() {
  // linkColumn is server-only: the browser needs UI validation metadata, not routing data.
  const attachments = {};
  Object.keys(EDIT_CONFIG.attachments).forEach(id => {
    const config = EDIT_CONFIG.attachments[id];
    attachments[id] = {
      title: config.title,
      enabled: config.enabled,
      allowedExtensions: config.allowedExtensions,
      maxSizeMb: config.maxSizeMb,
      replaceMode: config.replaceMode
    };
  });
  return { rules: EDIT_CONFIG.rules, fields: EDIT_CONFIG.fields, attachments };
}

function getMaxEnrollmentYear_() {
  return new Date().getFullYear() + 1;
}

function isEnrollmentYearInRange_(value) {
  const year = Number(value);
  return Number.isInteger(year) && year >= 1950 && year <= getMaxEnrollmentYear_();
}

function isValidRuDate_(value) {
  const match = String(value || '').match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!match) return false;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);

  if (month < 1 || month > 12) return false;
  if (day < 1) return false;

  const daysInMonth = new Date(year, month, 0).getDate();
  return day <= daysInMonth;
}

function validateEditableFieldValue_(columnName, value) {
  const fieldConfig = EDIT_CONFIG.fields[String(columnName || '')];
  if (!fieldConfig) throw new Error('Поле не настроено для редактирования.');
  if (fieldConfig.editable !== true) throw new Error('Изменение этого поля запрещено.');

  const rule = EDIT_CONFIG.rules[fieldConfig.rule];
  if (!rule) throw new Error('Для поля не найдено правило проверки.');

  const normalizedValue = String(value ?? '').trim();
  if (normalizedValue === '') {
    if (fieldConfig.required) throw new Error('Обязательное поле нельзя оставить пустым.');
    return { fieldConfig, rule, value: normalizedValue };
  }

  if (!new RegExp(rule.regex).test(normalizedValue)) {
    throw new Error(`Значение поля "${columnName}" не соответствует правилу: ${rule.title}.`);
  }

  if (rule.special === 'year' && !isEnrollmentYearInRange_(normalizedValue)) {
    throw new Error(`Год должен быть в диапазоне 1950-${getMaxEnrollmentYear_()}.`);
  }

  if (rule.special === 'date' && !isValidRuDate_(normalizedValue)) {
    throw new Error(`Значение поля "${columnName}" не является корректной календарной датой.`);
  }

  return { fieldConfig, rule, value: normalizedValue };
}

function prependSiteEditNote_(cell, historyTimestamp, newValue) {
  const note = cell.getNote() || '';
  const historyLine = `С: ${historyTimestamp}, ${newValue}`;
  cell.setNote(note ? `${historyLine}\n${note}` : historyLine);
}

function parseRuDateToDate_(value) {
  const match = String(value || '').match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);

  return new Date(year, month - 1, day);
}

function setValueWithSiteEditNote_(cell, newValue, historyTimestamp, options = {}) {
  const oldValue = formatCellValue(cell.getValue()).trim();
  if (oldValue === newValue) return false;

  let valueToSet = newValue;

  if (options.isDate) {
    valueToSet = parseRuDateToDate_(newValue);

    if (!valueToSet) {
      throw new Error('Не удалось преобразовать дату.');
    }
  }

  cell.setValue(valueToSet);
  prependSiteEditNote_(cell, historyTimestamp, newValue);

  return true;
}

function runFieldOnSaveAction_(fieldConfig, sheet, header, rowIndex, historyTimestamp) {
  if (fieldConfig.onSave !== 'UPDATE_SCHOOL_DATE') return false;
  const schoolUpdatedCol = header.indexOf('Дата обн. инф. о школе (С)');
  if (schoolUpdatedCol === -1) return false;
  const timestamp = historyTimestamp;
  const cell = sheet.getRange(rowIndex, schoolUpdatedCol + 1);
  return setValueWithSiteEditNote_(cell, timestamp, historyTimestamp);
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
  const validation = validateEditableFieldValue_(columnName, value);
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

  for (let i = 0; i < rowCount; i++) {
    const rowLogin = normalizeLogin(logins[i][0]);
    const rowPassword = formatCellValue(passwords[i][0]).trim();
    if (rowLogin !== normalizedLogin || rowPassword !== String(password || '').trim()) continue;

    const rowSnils = normalizeSnils(snilsValues[i][0]);
    if (rowSnils && normalizedSnils && rowSnils !== normalizedSnils) continue;

    const rowIndex = i + 2;

    const cell = sheet.getRange(rowIndex, targetCol + 1);
    const oldValue = formatCellValue(cell.getValue()).trim();
    const identityChanged = Boolean(validation.fieldConfig.isIdentityField && oldValue !== validation.value);
    const historyTimestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, `${CONFIG.DATE_FORMAT} HH:mm:ss`);
    const changed = setValueWithSiteEditNote_(cell, validation.value, historyTimestamp, {
  isDate: validation.rule.special === 'date'
});
    const onSaveChanged = changed ? runFieldOnSaveAction_(validation.fieldConfig, sheet, header, rowIndex, historyTimestamp) : false;
    return { ok: true, identityChanged: changed && identityChanged, changed, onSaveChanged };
  }

  throw new Error('Строка для обновления не найдена.');
}

/**
 * Uploads an attachment configured in EDIT_CONFIG. The client supplies no Drive,
 * spreadsheet or column identifiers; all sensitive routing is resolved here.
 */
function uploadAttachment(login, password, snils, attachmentId, filePayload) {
  try {
    const attachment = EDIT_CONFIG.attachments[String(attachmentId || '')];
    if (!attachment || attachment.enabled !== true) throw new Error('Этот тип документа отключён или настроен некорректно.');
    if (!filePayload || typeof filePayload !== 'object') throw new Error('Выберите файл для загрузки.');

    const fileName = String(filePayload.name || '');
    const extensionMatch = fileName.match(/\.([A-Za-z0-9]+)$/);
    const extension = extensionMatch ? extensionMatch[1].toLowerCase() : '';
    if (!extension || !attachment.allowedExtensions.includes(extension)) throw new Error('Недопустимое расширение файла.');
    const bytes = Utilities.base64Decode(String(filePayload.base64 || ''));
    const maxBytes = Number(attachment.maxSizeMb) * 1024 * 1024;
    if (!bytes.length || bytes.length > maxBytes) throw new Error(`Размер файла не должен превышать ${attachment.maxSizeMb} МБ.`);

    const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
    if (!sheet) throw new Error('Лист с результатами не найден.');
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) throw new Error('Таблица пуста.');
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const targetCol = header.indexOf(attachment.linkColumn);
    if (targetCol < 0) throw new Error('В таблице не найден столбец для ссылки на документ.');
    const rowIndex = findAuthorizedResultRow_(sheet, header, login, password, snils);
    const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const existingUrl = formatCellValue(row[targetCol]).trim();
    if (existingUrl && filePayload.confirmReplace !== true) {
      return { requiresConfirmation: true, currentFileUrl: existingUrl, currentFileName: getDriveFileNameFromUrl_(existingUrl) };
    }

    const folder = getOrCreateUserFolder_(header, row);
    const outputName = buildAttachmentFileName_(attachment.fileNameTemplate, header, row, extension);
    const blob = Utilities.newBlob(bytes, String(filePayload.mimeType || 'application/octet-stream'), outputName);
    let file;
    const oldFile = existingUrl ? getDriveFileFromUrl_(existingUrl) : null;
    if (oldFile && attachment.replaceMode === 'version') {
      oldFile.setContent(blob.getBytes());
      oldFile.setName(outputName);
      file = oldFile;
    } else {
      file = folder.createFile(blob);
      if (oldFile && attachment.replaceMode === 'replace') oldFile.setTrashed(true);
    }
    const url = file.getUrl();
    sheet.getRange(rowIndex, targetCol + 1).setValue(url);
    return { ok: true, url, fileName: file.getName(), replaced: Boolean(existingUrl) };
  } catch (error) {
    Logger.log(`[ATTACHMENT] ${error && error.stack ? error.stack : error}`);
    throw new Error(error && error.message ? error.message : 'Не удалось загрузить документ.');
  }
}

function findAuthorizedResultRow_(sheet, header, login, password, snils) {
  const authCols = getAuthColumnIndexes(header);
  const snilsCol = getSnilsColumnIndex(header);
  const rowCount = sheet.getLastRow() - 1;
  const auth = loadAuthColumns(sheet, rowCount, authCols, snilsCol);
  const requestedLogin = normalizeLogin(login);
  const requestedSnils = normalizeSnils(snils);
  for (let i = 0; i < rowCount; i++) {
    if (normalizeLogin(auth.logins[i][0]) !== requestedLogin || formatCellValue(auth.passwords[i][0]).trim() !== String(password || '').trim()) continue;
    const storedSnils = auth.snilsValues ? normalizeSnils(auth.snilsValues[i][0]) : '';
    if (storedSnils && storedSnils !== requestedSnils) continue;
    return i + 2;
  }
  throw new Error('Не удалось подтвердить текущую запись пользователя.');
}

function getOrCreateUserFolder_(header, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let root;
    try { root = DriveApp.getFolderById(SERVER_CONFIG.drive.usersRootFolderId); } catch (_error) { throw new Error('Не удалось открыть папку документов в Google Drive.'); }
    const name = renderTemplate_(SERVER_CONFIG.userFolderTemplate, header, row) || 'Пользователь без имени';
    const folders = root.getFoldersByName(name);
    return folders.hasNext() ? folders.next() : root.createFolder(name);
  } finally {
    lock.releaseLock();
  }
}

function buildAttachmentFileName_(template, header, row, extension) {
  const base = (renderTemplate_(template, header, row) || 'Документ').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
  return `${base || 'Документ'}.${extension}`;
}

function renderTemplate_(template, header, row) {
  const athleteName = formatCellValue(row[header.indexOf('Фамилия Имя Отчество (С)')]).trim().split(/\s+/);
  return String(template || '').replace(/\{([^}]+)\}/g, (_, title) => {
    const key = String(title).trim();
    const index = header.indexOf(key);
    if (index >= 0) return formatCellValue(row[index]).trim();
    const nameParts = { 'Фамилия': athleteName[0], 'Имя': athleteName[1], 'Отчество': athleteName.slice(2).join(' ') };
    if (nameParts[key]) return nameParts[key];
    if (key === 'Дата рождения (разряд)') {
      const dateIndex = header.indexOf('Дата рождения (С)');
      return dateIndex >= 0 ? formatCellValue(row[dateIndex]).trim() : '';
    }
    return '';
  }).replace(/\s+/g, ' ').trim();
}

function getDriveFileFromUrl_(url) {
  const match = String(url || '').match(/[-\w]{25,}/);
  if (!match) return null;
  try { return DriveApp.getFileById(match[0]); } catch (_error) { return null; }
}

function getDriveFileNameFromUrl_(url) {
  const file = getDriveFileFromUrl_(url);
  return file ? file.getName() : 'текущий файл';
}

/**
 * Возвращает уникальные варианты подсказок для настроенного поля.
 * Источник данных определяется конфигурацией конкретного поля, а не клиентом.
 * @param {string} columnName Заголовок редактируемого столбца.
 * @returns {string[]}
 */
function getFieldSuggestions(columnName) {
  const fieldConfig = EDIT_CONFIG.fields[String(columnName || '')];
  if (!fieldConfig || fieldConfig.rule !== 'SUGGEST_TEXT') return [];

  const suggestions = fieldConfig.suggestions;
  if (!suggestions || typeof suggestions !== 'object') return [];

  const { sourceSheet, sourceHeader, startRow } = suggestions;
  if (!sourceSheet || !sourceHeader || !Number.isFinite(Number(startRow)) || Number(startRow) < 1) return [];

  const sheet = getSheet(String(sourceSheet));
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const colIndex = header.indexOf(String(sourceHeader));
  if (colIndex === -1) return [];

  const fromRow = Math.max(2, Number(startRow));
  if (fromRow > lastRow) return [];

  const values = sheet.getRange(fromRow, colIndex + 1, lastRow - fromRow + 1, 1).getValues()
    .flat()
    .map(value => String(value || '').trim())
    .filter(Boolean);

  const uniq = [];
  const seen = new Set();
  values.forEach(value => {
    if (seen.has(value)) return;
    seen.add(value);
    uniq.push(value);
  });

  return uniq;
}
