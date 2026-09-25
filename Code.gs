
/**
 * Собирает публичную конфигурацию из Config-файлов для index.html.
 * @returns {string}
 */
function getClientConfigScript() {
  const responsiveConfig = getResponsiveClientConfig_();
  const clientConfig = {
    storageKeys: Object.assign({}, GENERAL_CLIENT_CONFIG.storageKeys, TechnicalConfig.storageKeys),
    timings: Object.assign({}, GENERAL_CLIENT_CONFIG.timings, TRAINING_CONFIG.client.timings, DOCUMENTS_CONFIG.client.timings),
    urls: Object.assign({}, GENERAL_CLIENT_CONFIG.urls, TRAINING_CONFIG.client.urls, {
      trainingSheet: TRAINING_CONFIG.spreadsheetUrl
    }),
    authFields: {
      login: CONFIG.AUTH.loginHeader,
      password: CONFIG.AUTH.passwordHeader,
      snils: CONFIG.AUTH.snilsHeader
    },
    ui: TRAINING_CONFIG.client.ui,
    validation: {
      enrollmentYearMin: EDIT_CONFIG.rules.YEAR.min,
      enrollmentYearOffset: EDIT_CONFIG.rules.YEAR.maxOffset
    },
    responsive: responsiveConfig,
    documentSections: DOCUMENTS_CONFIG.client.documentSections
  };
  return `window.CLIENT_CONFIG = ${JSON.stringify(clientConfig).replace(/</g, '\\u003c')};`;
}

/**
 * Проверяет и возвращает публичные настройки адаптивности.
 * @returns {{pageMaxWidth: number, compactModeWidth: number, compactMinViewportWidth: number, minimumInterfaceScale: number}}
 */
function getResponsiveClientConfig_() {
  const pageMaxWidth = RESPONSIVE_CONFIG.pageMaxWidth;
  const compactModeWidth = RESPONSIVE_CONFIG.compactModeWidth;
  const compactMinViewportWidth = RESPONSIVE_CONFIG.compactMinViewportWidth;
  const minimumInterfaceScale = RESPONSIVE_CONFIG.minimumInterfaceScale;

  if (!Number.isFinite(pageMaxWidth) || !Number.isFinite(compactModeWidth) ||
      !Number.isFinite(compactMinViewportWidth) || !Number.isFinite(minimumInterfaceScale) ||
      pageMaxWidth <= 0 || compactModeWidth <= 0 || compactMinViewportWidth <= 0) {
    throw new Error('Границы адаптивности должны быть положительными числами в пикселях, а minimumInterfaceScale — числом.');
  }
  if (compactModeWidth >= pageMaxWidth) {
    throw new Error('Некорректные настройки адаптивности: compactModeWidth должен быть меньше pageMaxWidth.');
  }
  if (compactMinViewportWidth >= compactModeWidth) {
    throw new Error('Некорректные настройки адаптивности: compactMinViewportWidth должен быть меньше compactModeWidth.');
  }
  if (minimumInterfaceScale <= 0 || minimumInterfaceScale > 1) {
    throw new Error('Некорректные настройки адаптивности: minimumInterfaceScale должен быть больше 0 и не больше 1.');
  }

  return { pageMaxWidth, compactModeWidth, compactMinViewportWidth, minimumInterfaceScale };
}

/*************************************************
 * ИНФРАСТРУКТУРА: ДОСТУП К ТАБЛИЦАМ
 *************************************************/
/** @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} */
function getSpreadsheet() {
  return SpreadsheetApp.openById(extractGoogleResourceId_(CONFIG.SPREADSHEET_URL));
}


/**
 * Извлекает идентификатор Google Sheets или Google Drive из полной ссылки.
 * @param {string} url Ссылка на ресурс Google.
 * @returns {string}
 */
function extractGoogleResourceId_(url) {
  const value = String(url || '').trim();
  const match = value.match(/(?:\/d\/|\/folders\/|[?&]id=)([-\w]{20,})/);
  if (!match) throw new Error('Укажите корректную ссылку на ресурс Google.');
  return match[1];
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
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_TITLE)
    // HtmlService добавляет метатеги в итоговую страницу Web App.
    // Это гарантирует CSS-ширину viewport устройства, а не виртуальные ~980 px.
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
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
 * ЛОГИЧЕСКИЕ СОБЫТИЯ ЖУРНАЛА
 *
 * Новый механизм пока не подключён к действиям сайта. Идентификатор события
 * хранится в заметке ячейки даты, а не в видимом столбце и не в номере строки:
 * заметка перемещается вместе со строкой при вставке строк сверху.
 *************************************************/
const LOGICAL_EVENT_NOTE_PREFIX = 'logical-event-id:';
const LOGICAL_EVENT_SESSION_CACHE_PREFIX = 'logical-event-session:';
const LOGICAL_EVENT_CACHE_TTL_SECONDS = 21600;
const LOGICAL_EVENT_TEST_SHEET_NAME = '_Тест логических событий';

/** Создаёт ключ текущего логического события для сессии страницы. */
function getLogicalEventSessionCacheKey_(sessionId) {
  return `${LOGICAL_EVENT_SESSION_CACHE_PREFIX}${String(sessionId || '').trim()}`;
}

/** Подготавливает лист с десятью видимыми столбцами нового журнала. */
function prepareLogicalEventSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, USER_LOG_COLUMNS.length).setNumberFormat('@').setValues([USER_LOG_COLUMNS]);
    return sheet;
  }

  const header = sheet.getRange(1, 1, 1, USER_LOG_COLUMNS.length).getValues()[0];
  if (USER_LOG_COLUMNS.some((title, index) => header[index] !== title)) {
    sheet.getRange(1, 1, 1, USER_LOG_COLUMNS.length).setNumberFormat('@').setValues([USER_LOG_COLUMNS]);
  }
  return sheet;
}

/**
 * Находит строку события по постоянному внутреннему ID.
 * Номер строки намеренно не кэшируется и не используется как идентификатор.
 */
function findLogicalEventRow_(sheet, eventId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const note = `${LOGICAL_EVENT_NOTE_PREFIX}${eventId}`;
  const notes = sheet.getRange(2, 1, lastRow - 1, 1).getNotes();
  const offset = notes.findIndex(row => row[0] === note);
  return offset === -1 ? -1 : offset + 2;
}

/** Добавляет текст в одну ячейку события, сохраняя ранее записанные значения. */
function appendLogicalEventCell_(sheet, rowIndex, columnIndex, value) {
  const cell = sheet.getRange(rowIndex, columnIndex);
  const previous = String(cell.getValue() || '');
  const next = String(value || '');
  cell.setNumberFormat('@').setValue(previous && next ? `${previous}\n${next}` : previous || next).setWrap(true);
}

/**
 * Создаёт самостоятельное событие и возвращает его постоянный внутренний ID.
 * @returns {{id:string, sessionId:string}}
 */
function createLogicalLogEvent_(sessionId, { login = '', password = '', snils = '', clientInfo = {} } = {}, targetSheet = null) {
  const normalizedSessionId = String(sessionId || '').trim();
  if (!normalizedSessionId) throw new Error('Для логического события требуется идентификатор сессии.');

  const sheet = prepareLogicalEventSheet_(targetSheet || getSheet(CONFIG.LOG_SHEET_NAME, true));
  const eventId = Utilities.getUuid();
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, USER_LOG_COLUMNS.length)
    .setNumberFormat('@')
    .setValues([[formatLogDateTime(new Date()), login, password, snils, clientInfo.ip || '', clientInfo.device || '', clientInfo.browser || '', '', '', '']])
    .setWrap(true);
  sheet.getRange(rowIndex, 1).setNote(`${LOGICAL_EVENT_NOTE_PREFIX}${eventId}`);
  CacheService.getScriptCache().put(getLogicalEventSessionCacheKey_(normalizedSessionId), eventId, LOGICAL_EVENT_CACHE_TTL_SECONDS);
  return { id: eventId, sessionId: normalizedSessionId };
}

/** Подставляет разрешённые параметры в шаблон события журнала. */
function renderUserLogEventTemplate_(eventId, parameters = {}) {
  const event = USER_LOG_EVENTS[eventId];
  if (!event || !event.enabled) throw new Error(`Событие журнала «${eventId}» не настроено.`);
  return event.template.replace(/\{([^{}]+)\}/g, (_, name) => String(parameters[name] ?? ''));
}

/** Дописывает результат авторизации к уже созданному нажатию «Войти». */
function appendLoginEventResult_(eventId, resultEventId, parameters = {}, sessionId = '') {
  if (!eventId) return;
  appendLogicalLogResult_(eventId, renderUserLogEventTemplate_(resultEventId, parameters));
  finishLogicalLogEvent_(sessionId, eventId);
}

/** Создаёт клиентское действие или дописывает локальные данные в указанное событие. */
function logClientLogicalEvent({ sessionId = '', eventId = '', eventName = '', parameters = {} } = {}) {
  const event = USER_LOG_EVENTS[eventName];
  if (!event || !event.enabled || event.source !== USER_LOG_EVENT_SOURCES.CLIENT) return '';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return '';
  try {
    const text = renderUserLogEventTemplate_(eventName, parameters);
    if (event.type === USER_LOG_EVENT_TYPES.LOCAL_DATA) {
      if (!eventId) return '';
      appendLogicalLogLocalData_(eventId, text);
      return eventId;
    }
    const logicalEvent = createLogicalLogEvent_(sessionId);
    writeLogicalLogAction_(logicalEvent.id, text);
    finishLogicalLogEvent_(sessionId, logicalEvent.id);
    return logicalEvent.id;
  } finally {
    lock.releaseLock();
  }
}

/** Записывает действие пользователя в созданное событие. */
function writeLogicalLogAction_(eventId, action, targetSheet = null) {
  appendLogicalLogEventValue_(eventId, 8, action, targetSheet);
}

/** Дописывает результат сервера в созданное событие. */
function appendLogicalLogResult_(eventId, result, targetSheet = null) {
  appendLogicalLogEventValue_(eventId, 9, result, targetSheet);
}

/** Дописывает локальные данные браузера в созданное событие. */
function appendLogicalLogLocalData_(eventId, localData, targetSheet = null) {
  appendLogicalLogEventValue_(eventId, 10, localData, targetSheet);
}

/** Находит событие по ID и дописывает значение в его видимую ячейку. */
function appendLogicalLogEventValue_(eventId, columnIndex, value, targetSheet = null) {
  const normalizedEventId = String(eventId || '').trim();
  if (!normalizedEventId) throw new Error('Не указан идентификатор логического события.');

  const sheet = prepareLogicalEventSheet_(targetSheet || getSheet(CONFIG.LOG_SHEET_NAME, true));
  const rowIndex = findLogicalEventRow_(sheet, normalizedEventId);
  if (rowIndex < 2) throw new Error('Логическое событие не найдено.');
  appendLogicalEventCell_(sheet, rowIndex, columnIndex, value);
}

/** Завершает событие, не затрагивая другие события той же сессии. */
function finishLogicalLogEvent_(sessionId, eventId) {
  const key = getLogicalEventSessionCacheKey_(sessionId);
  const cache = CacheService.getScriptCache();
  if (cache.get(key) === String(eventId || '').trim()) cache.remove(key);
}

/**
 * Техническая проверка механизма без подключения реальных действий сайта.
 * Создаёт отдельный лист, проверяет связность после вставки строки сверху и
 * удаляет тестовый лист после успешной проверки.
 */
function testLogicalLogEventMechanism() {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.insertSheet(`${LOGICAL_EVENT_TEST_SHEET_NAME}-${Utilities.getUuid()}`);

  try {
    prepareLogicalEventSheet_(sheet);
    const firstEvent = createLogicalLogEvent_('session-a', {}, sheet);
    writeLogicalLogAction_(firstEvent.id, 'Тестовое действие', sheet);
    appendLogicalLogResult_(firstEvent.id, 'Тестовый результат', sheet);
    appendLogicalLogLocalData_(firstEvent.id, 'Тестовые локальные данные', sheet);
    const firstRowBeforeInsert = findLogicalEventRow_(sheet, firstEvent.id);
    sheet.insertRowsBefore(2, 1);
    appendLogicalLogResult_(firstEvent.id, 'Результат после вставки строки', sheet);
    const firstRowAfterInsert = findLogicalEventRow_(sheet, firstEvent.id);

    const nextEvent = createLogicalLogEvent_('session-a', {}, sheet);
    writeLogicalLogAction_(nextEvent.id, 'Следующее действие', sheet);
    const otherSessionEvent = createLogicalLogEvent_('session-b', {}, sheet);
    writeLogicalLogAction_(otherSessionEvent.id, 'Действие другой сессии', sheet);
    finishLogicalLogEvent_('session-a', nextEvent.id);
    finishLogicalLogEvent_('session-b', otherSessionEvent.id);

    const firstValues = sheet.getRange(firstRowAfterInsert, 8, 1, 3).getValues()[0];
    if (firstRowAfterInsert === firstRowBeforeInsert || firstValues[0] !== 'Тестовое действие' || firstValues[1] !== 'Тестовый результат\nРезультат после вставки строки' || firstValues[2] !== 'Тестовые локальные данные') throw new Error('Не пройдена проверка дописывания в исходное событие.');
    if (nextEvent.id === firstEvent.id || otherSessionEvent.id === firstEvent.id || findLogicalEventRow_(sheet, nextEvent.id) === findLogicalEventRow_(sheet, otherSessionEvent.id)) throw new Error('Не пройдена проверка независимости событий и сессий.');

    return { passed: true, eventRows: 3, rowMovedAfterInsert: true, separateSessions: true };
  } finally {
    spreadsheet.deleteSheet(sheet);
  }
}

/*************************************************
 * ЛОГИРОВАНИЕ
 *************************************************/
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
  const rowValues = values.concat(Array(Math.max(0, LOG_COLUMNS.length - values.length)).fill(''));
  range.setNumberFormat('@').setValues([rowValues.slice(0, LOG_COLUMNS.length)]).setWrap(true);
}

/**
 * Разбирает дату первой строки журнала в настроенном формате отображения.
 * @param {*} value Значение ячейки с датой.
 * @returns {Date | null}
 */
function parseLogDateTime(value) {
  if (value instanceof Date) return value;

  const firstLine = splitLogLines(value)[0] || '';
  const format = `${CONFIG.DATE_FORMAT} HH:mm:ss`;
  const tokens = [];
  const pattern = format.replace(/yyyy|MM|dd|HH|mm|ss/g, token => {
    tokens.push(token);
    return token === 'yyyy' ? '(\\d{4})' : '(\\d{2})';
  }).replace(/[.]/g, '\\.');
  const match = firstLine.match(new RegExp(`^${pattern}$`));
  if (!match) return null;

  const values = {};
  tokens.forEach((token, index) => { values[token] = Number(match[index + 1]); });
  return new Date(values.yyyy, values.MM - 1, values.dd, values.HH, values.mm, values.ss);
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
 * Возвращает ключ временного сопоставления строки и идентификатора сессии.
 * Идентификатор создаётся браузером на каждую загрузку страницы и не хранится
 * в колонках листа, чтобы сохранить исходную структуру из восьми столбцов.
 */
function getLogSessionCacheKey_(sessionId) {
  return `access-log-session:${String(sessionId || '').trim()}`;
}

/** Ищет строку, созданную для текущей сессии браузера. */
function findSessionLogRow_(sessionId) {
  const value = CacheService.getScriptCache().get(getLogSessionCacheKey_(sessionId));
  const rowIndex = Number(value);
  return Number.isInteger(rowIndex) && rowIndex > 1 ? rowIndex : -1;
}

/** Запоминает строку сессии на срок, достаточный для активной работы страницы. */
function rememberSessionLogRow_(sessionId, rowIndex) {
  CacheService.getScriptCache().put(getLogSessionCacheKey_(sessionId), String(rowIndex), 21600);
}

/** Добавляет событие только в историю даты и статуса текущей строки. */
function appendLogLine(sheet, rowIndex, status) {
  const range = sheet.getRange(rowIndex, 1, 1, LOG_COLUMNS.length);
  const oldValues = range.getValues()[0];
  [
    { col: 0, value: formatLogDateTime(new Date()) },
    { col: 7, value: String(status || '') }
  ].forEach(({ col, value }) => {
    const oldText = formatLogCellValue(col, oldValues[col]);
    const lines = oldText === '' ? [value] : oldText.split('\n').concat([value]);
    range.getCell(1, col + 1).setRichTextValue(buildLogRichText(lines));
  });
  range.setWrap(true);
}

/** Создаёт строку только для нового открытия сайта. */
function logPageOpen({ login = '', password = '', snils = '', clientInfo = {} } = {}, sessionId) {
  if (!String(sessionId || '').trim()) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return;
  try {
    const sheet = getLogSheet();
    if (!sheet) return;
    const rowIndex = sheet.getLastRow() + 1;
    appendPlainLogRow(sheet, [
      formatLogDateTime(new Date()), login, password, snils,
      clientInfo.ip || '', clientInfo.device || '', clientInfo.browser || '', 'Зашел на сайт'
    ]);
    rememberSessionLogRow_(sessionId, rowIndex);
  } finally {
    lock.releaseLock();
  }
}

/** Добавляет событие в строку уже созданной сессии. */
function logSessionEvent(sessionId, status) {
  if (!String(sessionId || '').trim()) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return;
  try {
    const rowIndex = findSessionLogRow_(sessionId);
    if (rowIndex < 2) return;
    const sheet = getLogSheet();
    if (!sheet || rowIndex > sheet.getLastRow()) return;
    appendLogLine(sheet, rowIndex, status);
  } finally {
    lock.releaseLock();
  }
}

/** Логирует результат авторизации, но не silent-проверки. */
function logAuthAttempt(payload) {
  if (payload.clientInfo && payload.clientInfo.silent) return;
  logSessionEvent(payload.sessionId, payload.status);
}

function logUserAction(sessionId, status) {
  logSessionEvent(sessionId, status);
}

/**
 * Первый подключённый сценарий нового журнала: самостоятельное нажатие «Войти».
 * Его внутренний ID не передаётся в браузер и не виден в таблице.
 */
function logLoginButtonClick({ sessionId = '', login = '', password = '', snils = '', clientInfo = {} } = {}) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return;
  try {
    const event = createLogicalLogEvent_(sessionId, { login, password, snils, clientInfo });
    writeLogicalLogAction_(event.id, USER_LOG_EVENTS.login_click.template);
    return event.id;
  } finally {
    lock.releaseLock();
  }
}

function logSectionVisit(sessionId, section) {
  const sectionName = APP_CONFIG.SECTION_NAMES[section] || section;
  logSessionEvent(sessionId, `Перешёл в раздел ${sectionName}`);
}

function logFillFormClick(sessionId) {
  logSessionEvent(sessionId, 'Нажал: Заполнить форму');
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
  const loginCol = header.indexOf(CONFIG.AUTH.loginHeader);
  const passCol = header.indexOf(CONFIG.AUTH.passwordHeader);

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
  return header.indexOf(CONFIG.AUTH.snilsHeader);
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
 * Форматирует значение пароля для авторизации по правилу ввода DATE_RU.
 * Формат отображения CONFIG.DATE_FORMAT не влияет на вход в приложение.
 * @param {*} value Значение ячейки с датой рождения.
 * @returns {string}
 */
function formatAuthPasswordValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd.MM.yyyy');
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
    colors: allowedCols.map(i => backgrounds[i]),
    // FILE-поля хранят URL только на сервере. Клиент получает признак наличия,
    // а содержимое файла выдаётся отдельным авторизованным запросом.
    fileStates: allowedCols.map(i => {
      const fieldConfig = EDIT_CONFIG.fields[header[i]];
      return fieldConfig && fieldConfig.rule === 'FILE'
        ? { hasFile: Boolean(formatCellValue(row[i]).trim()) }
        : null;
    })
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
function checkLogin(login, password, clientInfo = {}, snils = '', sessionId = '', loginEventId = '') {
  const startedAt = Date.now();
  logStage('Начало checkLogin', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);

  if (!sheet) {
    appendLoginEventResult_(loginEventId, 'login_sheet_missing', {}, sessionId);
    return { error: 'Лист с результатами не найден.' };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  logStage(`Получены границы листа: rows=${lastRow}, cols=${lastCol}`, startedAt);

  if (lastRow < 2 || lastCol < 1) {
    appendLoginEventResult_(loginEventId, 'login_config_error', { message: 'Таблица пуста' }, sessionId);
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
    appendLoginEventResult_(loginEventId, 'login_config_error', { message: 'Ошибка структуры таблицы' }, sessionId);
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
    const rowPassword = formatAuthPasswordValue_(passwords[i][0]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = snilsValues ? normalizeSnils(snilsValues[i][0]) : '';
      if (rowSnils && rowSnils !== expectedSnils) {
        const snilsVisible = Boolean(clientInfo.snilsVisible);
        appendLoginEventResult_(loginEventId, expectedSnils && snilsVisible ? 'login_invalid_snils' : 'login_snils_required', {}, sessionId);
        logStage(expectedSnils && snilsVisible ? 'Совпадение найдено, СНИЛС неверный' : 'Совпадение найдено, требуется СНИЛС', startedAt);
        return { requiresSnils: true, snilsError: expectedSnils && snilsVisible ? 'invalid' : 'required' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      appendLoginEventResult_(loginEventId, 'login_success_without_snils', {}, sessionId);
      logStage('Совпадение найдено, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  appendLoginEventResult_(loginEventId, 'login_failed_credentials', {}, sessionId);
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
function verifySnils(login, password, snils, clientInfo = {}, sessionId = '', loginEventId = '') {
  const startedAt = Date.now();
  logStage('Начало verifySnils', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) {
    appendLoginEventResult_(loginEventId, 'login_sheet_missing', {}, sessionId);
    return { error: 'Лист с результатами не найден.' };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    appendLoginEventResult_(loginEventId, 'login_config_error', { message: 'Таблица пуста' }, sessionId);
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
    appendLoginEventResult_(loginEventId, 'login_config_error', { message: 'Ошибка структуры таблицы' }, sessionId);
    return { error: 'Ошибка структуры таблицы.' };
  }

  const snilsCol = getSnilsColumnIndex(header);
  if (snilsCol === -1) {
    appendLoginEventResult_(loginEventId, 'login_config_error', { message: 'Не настроен столбец СНИЛС' }, sessionId);
    return { error: 'Ошибка структуры таблицы.' };
  }

  const rowCount = lastRow - 1;
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, rowCount, authCols, snilsCol);
  logStage('Загружены колонки для проверки СНИЛС', startedAt);

  const normalizedLogin = normalizeLogin(login);
  const expectedSnils = normalizeSnils(snils);

  for (let i = 0; i < rowCount; i++) {
    const rowLogin = normalizeLogin(logins[i][0]);
    const rowPassword = formatAuthPasswordValue_(passwords[i][0]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = normalizeSnils(snilsValues[i][0]);
      if (!rowSnils) {
        const rowIndex = i + 2;
        const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
        const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
        appendLoginEventResult_(loginEventId, 'login_success_without_snils', {}, sessionId);
        return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
      }

      if (rowSnils !== expectedSnils) {
        appendLoginEventResult_(loginEventId, 'login_invalid_snils', {}, sessionId);
        return { error: 'Неверный СНИЛС.' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      appendLoginEventResult_(loginEventId, 'login_success_with_snils', {}, sessionId);
      logStage('СНИЛС подтвержден, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  appendLoginEventResult_(loginEventId, 'login_failed_credentials', {}, sessionId);
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
 * Возвращает ключ кэша ФИО с учётом настроенного столбца.
 * @returns {string}
 */
function getFullNamesCacheKey_() {
  return `${TechnicalConfig.NAMES_CACHE_KEY}:${TRAINING_CONFIG.athleteNameHeader}`;
}

/**
 * Загружает полные ФИО из настроенного столбца листа результатов.
 * Для ускорения используется краткоживущий кэш ScriptCache.
 * @returns {Array<{original:string,last:string,first:string,middle:string}>}
 */
function loadFullNames() {
  const cache = CacheService.getScriptCache();
  const cacheKey = getFullNamesCacheKey_();
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const fioCol = header.indexOf(TRAINING_CONFIG.athleteNameHeader);
  if (fioCol === -1) {
    throw new Error(`В листе «${CONFIG.RESULT_SHEET_NAME}» нет столбца ФИО «${TRAINING_CONFIG.athleteNameHeader}».`);
  }

  const values = sheet
    .getRange(2, fioCol + 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String)
    .map(normalizeFullName);

  cache.put(cacheKey, JSON.stringify(values), TRAINING_CONFIG.NAMES_CACHE_TTL_SECONDS);
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

  const maxErrors = TRAINING_CONFIG.maxNameMatchErrors;
  const matches = fullNames
    .map(full => calculateMatch(short, full, maxErrors))
    .filter(Boolean)
    .sort((a, b) => a.totalCost - b.totalCost);

  if (!matches.length) return null;

  const exactLast = matches.filter(m => m.lastCost === 0);
  const selected = exactLast.length === 1 ? exactLast[0].original : matches[0].original;

  const top = matches.slice(0, TRAINING_CONFIG.nameMatchOptionsLimit);
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
  const ss = SpreadsheetApp.openById(extractGoogleResourceId_(TRAINING_CONFIG.spreadsheetUrl));
  const sheet = ss.getSheetByName(TRAINING_CONFIG.responseSheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(String);
  const tsCol = header.indexOf(TRAINING_CONFIG.headers.timestamp);
  const dateCol = header.indexOf(TRAINING_CONFIG.headers.date);
  const coachCol = header.indexOf(TRAINING_CONFIG.headers.coach);
  const placeCol = header.indexOf(TRAINING_CONFIG.headers.place);

  if (tsCol === -1 || dateCol === -1 || coachCol === -1 || placeCol === -1) return [];

  const fioGroupMap = loadTrainingGroupMap();
  const startCol = TRAINING_CONFIG.responseFioStartColumn;
  const endCol = TRAINING_CONFIG.responseFioEndColumn;

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
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0].map(String);
  const fioCol = header.indexOf(TRAINING_CONFIG.athleteNameHeader);
  const groupCol = header.indexOf(TRAINING_CONFIG.athleteGroupHeader);
  if (fioCol === -1 || groupCol === -1) return {};

  const map = {};
  values.slice(1).forEach(row => {
    const fio = String(row[fioCol] || '').trim();
    const group = String(row[groupCol] || '').trim();
    if (!fio) return;
    map[normalizeTrainingName(fio)] = group || TRAINING_CONFIG.noGroupLabel;
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
    const group = fioGroupMap[normalized] || TRAINING_CONFIG.noGroupLabel;
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
    const pattern = hasTime ? `${CONFIG.DATE_FORMAT} HH:mm:ss` : CONFIG.DATE_FORMAT;
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
 * Возвращает браузеру конфигурацию редактирования из единственного
 * серверного источника истины. Клиент использует её только для UI/UX;
 * updateResultCell() всё равно проверяет права напрямую по EDIT_CONFIG.
 * @returns {{rules:Object, fields:Object}}
 */
function getEditConfig() {
  return EDIT_CONFIG;
}

function getMaxEnrollmentYear_() {
  return new Date().getFullYear() + EDIT_CONFIG.rules.YEAR.maxOffset;
}

function isEnrollmentYearInRange_(value) {
  const year = Number(value);
  return Number.isInteger(year) && year >= EDIT_CONFIG.rules.YEAR.min && year <= getMaxEnrollmentYear_();
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

  if (rule.special === 'file') {
    throw new Error('Файл можно изменить только через прикрепление.');
  }

  const normalizedValue = String(value ?? '').trim();
  if (normalizedValue === '') {
    if (fieldConfig.required) throw new Error('Обязательное поле нельзя оставить пустым.');
    return { fieldConfig, rule, value: normalizedValue };
  }

  if (!new RegExp(rule.regex).test(normalizedValue)) {
    throw new Error(`Значение поля "${columnName}" не соответствует правилу: ${rule.title}.`);
  }

  if (rule.special === 'year' && !isEnrollmentYearInRange_(normalizedValue)) {
    throw new Error(`Год должен быть в диапазоне ${EDIT_CONFIG.rules.YEAR.min}-${getMaxEnrollmentYear_()}.`);
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
  const schoolUpdatedCol = header.indexOf(DOCUMENTS_CONFIG.schoolInfoUpdatedHeader);
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
    const rowPassword = formatAuthPasswordValue_(passwords[i][0]).trim();
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
 * Подставляет значения указанной строки в шаблон имени.
 * Заголовки таблицы указываются в шаблоне маркерами `{Название столбца}`.
 * @param {{template:string}} namingConfig Конфигурация шаблона.
 * @param {string[]} header Заголовки листа «Результат».
 * @param {Array<*>} row Значения строки пользователя.
 * @returns {string}
 */
function renderAttachmentTemplate_(namingConfig, header, row) {
  if (!namingConfig || !String(namingConfig.template || '').trim()) {
    throw new Error('Шаблон имени вложения не настроен.');
  }

  const result = String(namingConfig.template).replace(/\{([^{}]*)\}/g, (_, marker) => {
    const columnName = String(marker).trim();
    if (!columnName) {
      throw new Error('В шаблоне имени указан пустой маркер столбца.');
    }
    const columnIndex = header.indexOf(columnName);
    if (columnIndex === -1) {
      throw new Error(`В листе «${CONFIG.RESULT_SHEET_NAME}» нет столбца «${columnName}» из шаблона.`);
    }
    const value = formatCellValue(row[columnIndex]).trim();
    // Пустая ячейка — допустимая часть имени. Это отличается от отсутствующего
    // заголовка, который проверяется выше и по-прежнему является ошибкой
    // настройки шаблона.
    return value;
  });

  return sanitizeDriveName_(result);
}

/**
 * Удаляет символы, недопустимые в имени объекта Google Drive, и лишние пробелы.
 * @param {string} value Исходное имя.
 * @returns {string}
 */
function sanitizeDriveName_(value) {
  const name = String(value || '').replace(/[\\/\u0000]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!name) throw new Error('После обработки шаблона получилось пустое имя.');
  return name;
}

/**
 * Возвращает расширение исходного файла, включая точку, либо пустую строку.
 * @param {string} fileName Имя исходного файла.
 * @returns {string}
 */
function getFileExtension_(fileName) {
  const baseName = String(fileName || '').trim();
  const dotIndex = baseName.lastIndexOf('.');
  return dotIndex > 0 && dotIndex < baseName.length - 1 ? baseName.slice(dotIndex) : '';
}

/**
 * Находит или создаёт папку пользователя в корневой папке вложений.
 * @param {GoogleAppsScript.Drive.Folder} rootFolder Корневая папка «Пользователи».
 * @param {string} folderName Имя папки пользователя.
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateUserAttachmentsFolder_(rootFolder, folderName) {
  const folders = rootFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
}

/**
 * Перемещает заменяемый файл приложения в корзину, если в ячейке есть ссылка Drive.
 * @param {*} value URL прежнего файла.
 */
function movePreviousAttachmentToTrash_(value) {
  const url = String(value || '');
  const match = url.match(/(?:\/d\/|[?&]id=)([-\w]{20,})/);
  if (!match) return;
  try {
    DriveApp.getFileById(match[1]).setTrashed(true);
  } catch (error) {
    // Старый файл уже мог быть удалён или быть недоступным; новый файл остаётся валидным.
    Logger.log(`Не удалось переместить прежнее вложение в корзину: ${error.message}`);
  }
}

/**
 * Загружает файл для FILE-поля и сохраняет его URL в той же ячейке.
 * Это делает состояние файла частью данных строки; обычное текстовое
 * сохранение после загрузки не требуется.
 * @param {{login:string,password:string,snils:string,columnName:string,attachmentFile:GoogleAppsScript.Base.Blob}} formData
 * @returns {{ok:boolean,fileName:string,fileState:{hasFile:boolean}}}
 */
function uploadDocumentAttachment(formData) {
  const payload = formData || {};
  const columnName = String(payload.columnName || '');
  const fieldConfig = EDIT_CONFIG.fields[columnName];
  if (!fieldConfig || fieldConfig.rule !== 'FILE' || fieldConfig.editable !== true) {
    throw new Error('Для этого поля прикрепление файла не настроено.');
  }

  const file = payload.attachmentFile;
  if (!file || typeof file.getBytes !== 'function' || !file.getBytes().length) {
    throw new Error('Выберите файл для загрузки.');
  }

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) throw new Error('Лист с результатами не найден.');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('Таблица пуста.');

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const targetCol = header.indexOf(columnName);
  if (targetCol === -1) throw new Error('Колонка не найдена.');

  const authCols = getAuthColumnIndexes(header);
  const snilsCol = getSnilsColumnIndex(header);
  const rowCount = lastRow - 1;
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, rowCount, authCols, snilsCol);
  const normalizedLogin = normalizeLogin(payload.login);
  const password = String(payload.password || '').trim();
  const normalizedSnils = normalizeSnils(payload.snils);

  for (let i = 0; i < rowCount; i++) {
    if (normalizeLogin(logins[i][0]) !== normalizedLogin || formatAuthPasswordValue_(passwords[i][0]).trim() !== password) continue;
    const rowSnils = snilsValues ? normalizeSnils(snilsValues[i][0]) : '';
    if (rowSnils && rowSnils !== normalizedSnils) continue;

    const row = sheet.getRange(i + 2, 1, 1, lastCol).getValues()[0];
    const folderName = renderAttachmentTemplate_(DOCUMENTS_CONFIG.userFolder, header, row);
    const fileTemplate = DOCUMENTS_CONFIG.files[columnName];
    if (!fileTemplate) throw new Error(`Для FILE-поля «${columnName}» не настроен шаблон имени.`);

    const fileBaseName = renderAttachmentTemplate_(fileTemplate, header, row);
    const extension = getFileExtension_(file.getName());
    const destinationFolder = getOrCreateUserAttachmentsFolder_(
      DriveApp.getFolderById(extractGoogleResourceId_(DOCUMENTS_CONFIG.attachmentsFolderUrl)),
      folderName
    );
    const uploadedFile = destinationFolder.createFile(file).setName(`${fileBaseName}${extension}`);
    movePreviousAttachmentToTrash_(row[targetCol]);
    const historyTimestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, `${CONFIG.DATE_FORMAT} HH:mm:ss`);
    setValueWithSiteEditNote_(sheet.getRange(i + 2, targetCol + 1), uploadedFile.getUrl(), historyTimestamp);
    return { ok: true, fileName: uploadedFile.getName(), fileState: { hasFile: true } };
  }

  throw new Error('Не удалось подтвердить пользователя для загрузки файла.');
}

/**
 * Однократно отзывает общий доступ по ссылке у уже загруженных FILE-вложений.
 * Запустите вручную из редактора Apps Script после развёртывания изменения.
 * @returns {{checked:number,revoked:number,errors:number}}
 */
function revokeExistingAttachmentLinkSharing() {
  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { checked: 0, revoked: 0, errors: 0 };
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const fileColumns = header.map((name, index) => EDIT_CONFIG.fields[name]?.rule === 'FILE' ? index : -1).filter(index => index >= 0);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const result = { checked: 0, revoked: 0, errors: 0 };
  values.forEach(row => fileColumns.forEach(columnIndex => {
    const match = formatCellValue(row[columnIndex]).match(/(?:\/d\/|[?&]id=)([-\w]{20,})/);
    if (!match) return;
    result.checked++;
    try {
      DriveApp.getFileById(match[1]).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      result.revoked++;
    } catch (error) {
      result.errors++;
      Logger.log(`Не удалось отозвать доступ по ссылке к вложению: ${error.message}`);
    }
  }));
  return result;
}

/**
 * Возвращает содержимое прикреплённого файла только после проверки пользователя.
 * URL Drive не передаётся браузеру, поэтому доступ по ссылке не требуется.
 * @param {string} login
 * @param {string} password
 * @param {string} snils
 * @param {string} columnName
 * @returns {{fileName:string,mimeType:string,base64:string}}
 */
function getDocumentAttachmentContent(login, password, snils, columnName) {
  const fieldConfig = EDIT_CONFIG.fields[String(columnName || '')];
  if (!fieldConfig || fieldConfig.rule !== 'FILE') throw new Error('Для этого поля нет прикреплённого файла.');
  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) throw new Error('Лист с результатами не найден.');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('Таблица пуста.');
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const targetCol = header.indexOf(String(columnName));
  if (targetCol === -1) throw new Error('Колонка не найдена.');
  const authCols = getAuthColumnIndexes(header);
  const snilsCol = getSnilsColumnIndex(header);
  const { logins, passwords, snilsValues } = loadAuthColumns(sheet, lastRow - 1, authCols, snilsCol);
  const normalizedLogin = normalizeLogin(login);
  const normalizedSnils = normalizeSnils(snils);
  for (let i = 0; i < lastRow - 1; i++) {
    if (normalizeLogin(logins[i][0]) !== normalizedLogin || formatAuthPasswordValue_(passwords[i][0]).trim() !== String(password || '').trim()) continue;
    const rowSnils = snilsValues ? normalizeSnils(snilsValues[i][0]) : '';
    if (rowSnils && rowSnils !== normalizedSnils) continue;
    const fileUrl = formatCellValue(sheet.getRange(i + 2, targetCol + 1).getValue()).trim();
    const match = fileUrl.match(/(?:\/d\/|[?&]id=)([-\w]{20,})/);
    if (!match) throw new Error('Прикреплённый файл не найден.');
    const file = DriveApp.getFileById(match[1]);
    const blob = file.getBlob();
    return { fileName: file.getName(), mimeType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
  }
  throw new Error('Не удалось подтвердить пользователя для открытия файла.');
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
