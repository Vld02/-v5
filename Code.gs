
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
  range.setNumberFormat('@').setValues([values]).setWrap(true);
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

/**
 * Добавляет скрытые метаданные к строке логического события.
 * Метаданные диапазона остаются связанными со строкой после вставки строк
 * выше неё, поэтому номер строки нигде не выступает идентификатором события.
 */
function addLogicalEventMetadata_(range, key, value) {
  range.addDeveloperMetadata(key, String(value), SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
}

/** Находит диапазон строки по постоянному внутреннему ID логического события. */
function findLogicalEventRange_(sheet, eventId) {
  const metadata = sheet.createDeveloperMetadataFinder()
    .withKey(LOGICAL_EVENT_LOG_CONFIG.eventIdMetadataKey)
    .withValue(String(eventId || ''))
    .find();
  if (metadata.length !== 1) return null;

  const location = metadata[0].getLocation();
  return location && location.getRange ? location.getRange() : null;
}

/** Проверяет, было ли логическое событие завершено. */
function isLogicalEventCompleted_(range, eventId) {
  return range.getDeveloperMetadata().some(metadata =>
    metadata.getKey() === LOGICAL_EVENT_LOG_CONFIG.stateMetadataKey &&
    metadata.getValue() === `${eventId}:${LOGICAL_EVENT_LOG_CONFIG.completedState}`
  );
}

/**
 * Создаёт независимое логическое событие и возвращает его постоянный ID.
 * Созданная строка содержит действие; ID и ID сессии существуют только в
 * метаданных диапазона и не добавляются в видимые столбцы журнала.
 *
 * @param {string} sessionId ID сессии страницы.
 * @param {string} action Текст самостоятельного пользовательского действия.
 * @returns {string} Внутренний ID события или пустая строка при ошибке записи.
 */
function createLogicalLogEvent(sessionId, action) {
  if (!String(sessionId || '').trim() || !String(action || '').trim()) return '';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return '';
  try {
    const sheet = getLogSheet();
    if (!sheet) return '';

    const eventId = Utilities.getUuid();
    const rowIndex = sheet.getLastRow() + 1;
    appendPlainLogRow(sheet, [
      formatLogDateTime(new Date()), '', '', '', '', '', '', String(action).trim(), '', ''
    ]);
    const range = sheet.getRange(rowIndex, 1, 1, LOG_COLUMNS.length);
    addLogicalEventMetadata_(range, LOGICAL_EVENT_LOG_CONFIG.eventIdMetadataKey, eventId);
    addLogicalEventMetadata_(range, LOGICAL_EVENT_LOG_CONFIG.sessionIdMetadataKey, String(sessionId).trim());
    addLogicalEventMetadata_(range, LOGICAL_EVENT_LOG_CONFIG.stateMetadataKey, `${eventId}:active`);
    return eventId;
  } finally {
    lock.releaseLock();
  }
}

/** Дописывает результат в ранее созданное логическое событие. */
function appendLogicalLogResult(eventId, result) {
  return appendLogicalLogPart_(eventId, 'Результат', result);
}

/** Дописывает локальные данные в ранее созданное логическое событие. */
function appendLogicalLogLocalData(eventId, localData) {
  return appendLogicalLogPart_(eventId, 'Локальные данные', localData);
}

/** Находит событие по ID и добавляет к нему часть без использования номера строки как ID. */
function appendLogicalLogPart_(eventId, label, value) {
  if (!String(eventId || '').trim() || !String(value || '').trim()) return false;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return false;
  try {
    const sheet = getLogSheet();
    const range = sheet && findLogicalEventRange_(sheet, eventId);
    if (!range || isLogicalEventCompleted_(range, eventId)) return false;
    const column = label === 'Результат' ? 9 : 10;
    const cell = range.getCell(1, column);
    const previous = String(cell.getValue() || '');
    cell.setValue(previous ? `${previous}\n${String(value).trim()}` : String(value).trim()).setWrap(true);
    return true;
  } finally { lock.releaseLock(); }
}

/** Завершает логическое событие и запрещает последующие дописывания в него. */
function completeLogicalLogEvent(eventId) {
  if (!String(eventId || '').trim()) return false;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOG_CONFIG.lockWaitMs)) return false;
  try {
    const sheet = getLogSheet();
    if (!sheet) return false;
    const range = findLogicalEventRange_(sheet, eventId);
    if (!range || isLogicalEventCompleted_(range, eventId)) return false;
    range.getDeveloperMetadata().filter(metadata =>
      metadata.getKey() === LOGICAL_EVENT_LOG_CONFIG.stateMetadataKey
    ).forEach(metadata => metadata.remove());
    addLogicalEventMetadata_(
      range,
      LOGICAL_EVENT_LOG_CONFIG.stateMetadataKey,
      `${eventId}:${LOGICAL_EVENT_LOG_CONFIG.completedState}`
    );
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Подставляет переданные параметры в шаблон события из единой конфигурации. */
function renderLogEventTemplate_(event, params = {}) {
  return event.template.replace(/\{([A-Za-z][A-Za-z0-9]*)\}/g, (_match, name) =>
    String(params[name] ?? '')
  );
}

/** Возвращает включённое событие ожидаемого типа из конфигурации. */
function getConfiguredLogEvent_(eventId, expectedType) {
  const event = LOG_EVENT_CONFIG.EVENTS[eventId];
  if (!event || !event.enabled || event.type !== expectedType) return null;
  return event;
}

/**
 * Создаёт действие по его ID из Config.gs. Клиент передаёт только ID и
 * параметры, а текст события остаётся единственным источником в конфигурации.
 */
function createConfiguredLogicalEvent(sessionId, eventId, params = {}) {
  const event = getConfiguredLogEvent_(eventId, LOG_EVENT_CONFIG.EVENT_TYPES.ACTION);
  if (!event || event.mode !== LOG_EVENT_CONFIG.RULES.CREATE_NEW) return '';
  const logicalEventId = createLogicalLogEvent(sessionId, renderLogEventTemplate_(event, params));
  const sheet = logicalEventId && getLogSheet();
  const range = sheet && findLogicalEventRange_(sheet, logicalEventId);
  if (range) { addLogicalEventMetadata_(range, 'logical-log-action-id', event.id); if (event.completeAfterWrite) completeLogicalLogEvent(logicalEventId); }
  return logicalEventId;
}

/** Дописывает результат или локальные данные по ID события из Config.gs. */
function appendConfiguredLogicalEventPart(eventId, partEventId, params = {}) {
  const event = LOG_EVENT_CONFIG.EVENTS[partEventId];
  if (!event || !event.enabled || event.mode !== LOG_EVENT_CONFIG.RULES.ATTACH_TO_EVENT || !event.parentEvent) return false;
  const sheet = getLogSheet();
  const range = sheet && findLogicalEventRange_(sheet, eventId);
  if (!range || isLogicalEventCompleted_(range, eventId)) return false;
  const session = range.getDeveloperMetadata().find(item => item.getKey() === LOGICAL_EVENT_LOG_CONFIG.sessionIdMetadataKey);
  const parent = range.getDeveloperMetadata().find(item => item.getKey() === 'logical-log-action-id');
  // Both IDs must be present; the requested child must belong to this exact action,
  // never just to the last action of the same kind.
  if (!session || !session.getValue() || !parent || parent.getValue() !== event.parentEvent) return false;
  const value = renderLogEventTemplate_(event, params);
  const written = event.type === LOG_EVENT_CONFIG.EVENT_TYPES.RESULT
    ? appendLogicalLogResult(eventId, value)
    : event.type === LOG_EVENT_CONFIG.EVENT_TYPES.LOCAL_DATA
      ? appendLogicalLogLocalData(eventId, value)
      : false;
  if (written && event.completeAfterWrite) completeLogicalLogEvent(eventId);
  return written;
}

/**
 * Ручной технический тест этапа 2. Не вызывается сайтом и не подключён к
 * пользовательским действиям. Он создаёт два события первой сессии и одно
 * второй, вставляет строку над первым и проверяет связь по метаданным.
 */
function runLogicalLogEventTechnicalTest() {
  const suffix = Utilities.getUuid();
  const firstSessionId = `logical-log-test-session-a-${suffix}`;
  const secondSessionId = `logical-log-test-session-b-${suffix}`;
  const firstEventId = createLogicalLogEvent(firstSessionId, 'Техническое действие');
  if (!firstEventId || !appendLogicalLogResult(firstEventId, 'Технический результат')) {
    throw new Error('Не удалось создать событие или дописать результат.');
  }

  const sheet = getLogSheet();
  const firstRangeBeforeInsert = findLogicalEventRange_(sheet, firstEventId);
  if (!firstRangeBeforeInsert) throw new Error('Не найдено созданное событие.');
  sheet.insertRowBefore(firstRangeBeforeInsert.getRow());

  if (!appendLogicalLogLocalData(firstEventId, 'Технические локальные данные')) {
    throw new Error('Вставка строки нарушила связь с событием.');
  }
  if (!completeLogicalLogEvent(firstEventId) || appendLogicalLogResult(firstEventId, 'Не должно быть записано')) {
    throw new Error('Завершение события работает неверно.');
  }

  const nextEventId = createLogicalLogEvent(firstSessionId, 'Следующее техническое действие');
  const secondSessionEventId = createLogicalLogEvent(secondSessionId, 'Действие второй сессии');
  if (!nextEventId || !secondSessionEventId || nextEventId === firstEventId || secondSessionEventId === firstEventId) {
    throw new Error('Новые события не получили независимые ID.');
  }

  const firstRange = findLogicalEventRange_(sheet, firstEventId);
  const nextRange = findLogicalEventRange_(sheet, nextEventId);
  const secondSessionRange = findLogicalEventRange_(sheet, secondSessionEventId);
  const firstAction = String(firstRange.getCell(1, 8).getValue());
  const firstResult = String(firstRange.getCell(1, 9).getValue());
  const firstLocalData = String(firstRange.getCell(1, 10).getValue());
  const firstSessionMetadata = firstRange.getDeveloperMetadata().find(metadata =>
    metadata.getKey() === LOGICAL_EVENT_LOG_CONFIG.sessionIdMetadataKey
  );
  const secondSessionMetadata = secondSessionRange.getDeveloperMetadata().find(metadata =>
    metadata.getKey() === LOGICAL_EVENT_LOG_CONFIG.sessionIdMetadataKey
  );
  if (
    firstAction !== 'Техническое действие' ||
    firstResult !== 'Технический результат' ||
    firstLocalData !== 'Технические локальные данные' ||
    !firstSessionMetadata || firstSessionMetadata.getValue() !== firstSessionId ||
    !secondSessionMetadata || secondSessionMetadata.getValue() !== secondSessionId ||
    firstRange.getRow() === nextRange.getRow() || firstRange.getRow() === secondSessionRange.getRow() ||
    nextRange.getRow() === secondSessionRange.getRow()
  ) {
    throw new Error('Технический тест логических событий завершился с неверными данными.');
  }

  return {
    firstEventId,
    firstEventRow: firstRange.getRow(),
    nextEventId,
    nextEventRow: nextRange.getRow(),
    secondSessionEventId,
    secondSessionEventRow: secondSessionRange.getRow()
  };
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
function checkLogin(login, password, clientInfo = {}, snils = '', sessionId = '', logicalEventId = '') {
  const startedAt = Date.now();
  logStage('Начало checkLogin', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);

  if (!sheet) {
    logAuthResult_(logicalEventId, 'login_sheet_missing', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Лист не найден' });
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
    logAuthResult_(logicalEventId, 'login_config_error', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Ошибка конфигурации столбцов' });
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
        const snilsStatus = expectedSnils && snilsVisible ? 'Неверный СНИЛС' : 'Требуется ввод СНИЛС';
        const snilsEventId = expectedSnils && snilsVisible ? 'login_invalid_snils' : 'login_snils_required';
        logAuthResult_(logicalEventId, snilsEventId, clientInfo, { login, password, snils, clientInfo, sessionId, status: snilsStatus });
        logStage(expectedSnils && snilsVisible ? 'Совпадение найдено, СНИЛС неверный' : 'Совпадение найдено, требуется СНИЛС', startedAt);
        return { requiresSnils: true, snilsError: expectedSnils && snilsVisible ? 'invalid' : 'required' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      const successEventId = rowSnils ? 'login_success_with_snils' : 'login_success_without_snils';
      const successStatus = rowSnils ? 'Удачный вход по СНИЛС' : 'Удачный вход без СНИЛС';
      logAuthResult_(logicalEventId, successEventId, clientInfo, { login, password, snils, clientInfo, sessionId, status: successStatus });
      logStage('Совпадение найдено, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  logAuthResult_(logicalEventId, 'login_failed_credentials', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Неудачный вход: ФИО/дата' });
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
function verifySnils(login, password, snils, clientInfo = {}, sessionId = '', logicalEventId = '') {
  const startedAt = Date.now();
  logStage('Начало verifySnils', startedAt);

  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) {
    logAuthResult_(logicalEventId, 'login_sheet_missing', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Лист не найден' });
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
    logAuthResult_(logicalEventId, 'login_config_error', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Ошибка конфигурации столбцов' });
    return { error: 'Ошибка структуры таблицы.' };
  }

  const snilsCol = getSnilsColumnIndex(header);
  if (snilsCol === -1) {
    logAuthResult_(logicalEventId, 'login_config_error', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Ошибка конфигурации столбцов' });
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
        logAuthResult_(logicalEventId, 'login_success_without_snils', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Удачный вход без СНИЛС' });
        return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
      }

      if (rowSnils !== expectedSnils) {
        logAuthResult_(logicalEventId, 'login_invalid_snils', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Неверный СНИЛС' });
        return { error: 'Неверный СНИЛС.' };
      }

      const rowIndex = i + 2;
      const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
      const rowBackgrounds = sheet.getRange(rowIndex, 1, 1, lastCol).getBackgrounds()[0];
      logAuthResult_(logicalEventId, 'login_success_with_snils', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Удачный вход по СНИЛС' });
      logStage('СНИЛС подтвержден, данные строки загружены', startedAt);
      return prepareRowForClient(row, header, rowBackgrounds, allowedCols);
    }
  }

  logAuthResult_(logicalEventId, 'login_failed_credentials', clientInfo, { login, password, snils, clientInfo, sessionId, status: 'Неудачный вход: ФИО/дата' });
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
function updateResultCell(login, password, snils, columnName, value, logicalEventId = '') {
  try {
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
      appendSaveResult_(logicalEventId, 'save_success', columnName, validation.value);
      return { ok: true, identityChanged: changed && identityChanged, changed, onSaveChanged };
    }

    throw new Error('Строка для обновления не найдена.');
  } catch (error) {
    appendSaveResult_(logicalEventId, 'save_error', columnName);
    throw error;
  }
}

/** Дописывает итог серверного сохранения к уже созданному действию «Сохранить». */
function appendSaveResult_(logicalEventId, eventId, columnName, value = '') {
  if (!String(logicalEventId || '').trim()) return;
  try {
    const params = { field: String(columnName || '').replace(/\s*\([^()]+\)$/, '').trim() };
    if (eventId === 'save_success') params.value = value;
    appendConfiguredLogicalEventPart(logicalEventId, eventId, params);
  } catch (error) {
    Logger.log(`Не удалось записать результат сохранения: ${error.message}`);
  }
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
