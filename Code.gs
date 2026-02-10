/*************************************************
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 *************************************************/
const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ',
  RESULT_SHEET_NAME: 'Результат',
  LOG_SHEET_NAME: 'Входы',
  FORM_RESPONSES_SHEET_NAME: 'Ответы',
  FORM_USER_ID_HEADER: 'Фамилия Имя Отчество',
  FORM_SESSION_KEY_PREFIX: 'form_session_',
  GOOGLE_FORM_VIEW_URL: 'https://docs.google.com/forms/d/e/1FAIpQLSfLOcMYEWaOwF3qYuzsnSmaDbOSR6WGB7AUP7oYHBpzlo7npQ/viewform',
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
function doGet(e) {
  const fio = String((e && e.parameter && e.parameter.fio) || '').trim();
  if (fio) {
    const target = String((e && e.parameter && e.parameter.target) || '').trim();
    const redirectUrl = target || CONFIG.GOOGLE_FORM_VIEW_URL;
    return renderFormGatewayPage(fio, redirectUrl);
  }

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('ДБВv5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Возвращает HTML-контент файла.
 * @param {string} name Имя HTML-файла.
 * @returns {string}
 */
function getHtmlFile(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}


/**
 * Возвращает URL текущего Web App.
 * @returns {{url:string}}
 */
function getWebAppUrl() {
  return { url: ScriptApp.getService().getUrl() || '' };
}

/**
 * Экранирует HTML-строку.
 * @param {string} value
 * @returns {string}
 */
function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Сохраняет связку session_token -> ФИО.
 * @param {string} fio
 * @returns {string}
 */
function saveGatewaySession(fio) {
  const token = Utilities.getUuid();
  const payload = {
    fio: String(fio || '').trim(),
    createdAt: Date.now()
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(`${CONFIG.FORM_SESSION_KEY_PREFIX}${token}`, JSON.stringify(payload));

  return token;
}

/**
 * Рендерит страницу-прокладку перед переходом в форму.
 * @param {string} fio
 * @param {string} redirectUrl
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function renderFormGatewayPage(fio, redirectUrl) {
  const cleanFio = String(fio || '').trim();
  if (!cleanFio) {
    return HtmlService.createHtmlOutput('<h3>Не указан пользователь.</h3>');
  }

  const requestedTarget = String(redirectUrl || '').trim();
  const target = requestedTarget.startsWith('https://docs.google.com/forms/')
    ? requestedTarget
    : CONFIG.GOOGLE_FORM_VIEW_URL;
  saveGatewaySession(cleanFio);

  const html = `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Подтверждение пользователя</title>
  <style>
    body { font-family: Arial, sans-serif; background:#f5f5f5; margin:0; }
    .card { max-width:640px; margin:60px auto; background:#fff; border-radius:12px; padding:24px; box-shadow:0 8px 24px rgba(0,0,0,.08); }
    .fio { font-size:22px; font-weight:700; margin:10px 0 18px; }
    .warn { color:#a40000; margin:0 0 18px; }
    button { background:#7f0000; color:#fff; border:none; border-radius:8px; padding:12px 18px; cursor:pointer; font-size:16px; }
  </style>
</head>
<body>
  <div class="card">
    <div>Вы заполняете форму за:</div>
    <div class="fio">${escapeHtml(cleanFio)}</div>
    <p class="warn">Если это не вы — закройте страницу.</p>
    <button type="button" onclick="window.location.href='${escapeHtml(target)}'">Перейти к форме</button>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Подтверждение пользователя')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*************************************************
 * ЛОГИРОВАНИЕ
 *************************************************/
/**
 * Пишет запись об авторизации/действии в лист логов.
 * @param {{login?:string,password?:string,clientInfo?:Object,status:string}} payload
 */
function logAccess({ login = '', password = '', clientInfo = {}, status }) {
  const sheet = getSheet(CONFIG.LOG_SHEET_NAME, true);
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Дата/время',
      'Логин',
      'Пароль',
      'IP',
      'Устройство',
      'Браузер',
      'Статус'
    ]);
  }

  sheet.appendRow([
    new Date(),
    login,
    password,
    clientInfo.ip || '',
    clientInfo.device || '',
    clientInfo.browser || '',
    status
  ]);
}

/**
 * Логирует клик по кнопке открытия формы.
 * @param {string} login Логин пользователя.
 */
function logFormClick(login) {
  logAccess({ login, status: 'Нажал: Заполнить форму' });
}

/**
 * Извлекает и удаляет ближайшую сессию Web App для указанного времени отправки.
 * @param {Date} submittedAt
 * @returns {string}
 */
function consumeUserIdForSubmission(submittedAt) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const all = properties.getProperties();
    const submitTs = submittedAt instanceof Date ? submittedAt.getTime() : Date.now();

    let selectedKey = '';
    let selectedDelta = Number.POSITIVE_INFINITY;

    Object.keys(all).forEach(key => {
      if (!key.startsWith(CONFIG.FORM_SESSION_KEY_PREFIX)) return;

      let payload;
      try {
        payload = JSON.parse(all[key]);
      } catch (_error) {
        return;
      }

      const fio = String(payload.fio || '').trim();
      const createdAt = Number(payload.createdAt || 0);
      if (!fio || !createdAt) return;

      const delta = Math.abs(submitTs - createdAt);
      if (delta < selectedDelta) {
        selectedDelta = delta;
        selectedKey = key;
      }
    });

    if (!selectedKey) return '';

    const payload = JSON.parse(all[selectedKey]);
    properties.deleteProperty(selectedKey);
    return String(payload.fio || '').trim();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Заполняет в листе "Ответы" колонку "Фамилия Имя Отчество" после отправки Google Form.
 * Должно вызываться installable-триггером формы "On form submit".
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function onFormSubmit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== CONFIG.FORM_RESPONSES_SHEET_NAME) return;

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idCol = header.indexOf(CONFIG.FORM_USER_ID_HEADER);
  if (idCol === -1) {
    console.error('COLUMN_NOT_FOUND: Фамилия Имя Отчество');
    return;
  }

  const rowIndex = e.range.getRow();
  const currentValue = String(sheet.getRange(rowIndex, idCol + 1).getValue() || '').trim();
  if (currentValue) return;

  const submitDate = sheet.getRange(rowIndex, 1).getValue();
  const fio = consumeUserIdForSubmission(submitDate instanceof Date ? submitDate : new Date());
  if (!fio) {
    console.error(`FIO_NOT_RESOLVED_FOR_ROW: ${rowIndex}`);
    return;
  }

  sheet.getRange(rowIndex, idCol + 1).setValue(fio);
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
 * Проверяет логин/дату рождения и возвращает персональные данные для UI.
 * @param {string} login ФИО.
 * @param {string} password Дата рождения в формате ДД.ММ.ГГГГ.
 * @param {Object} [clientInfo={}] Данные об устройстве.
 * @returns {Object}
 */
function checkLogin(login, password, clientInfo = {}) {
  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);

  if (!sheet) {
    logAccess({ login, password, clientInfo, status: 'Лист не найден' });
    return { error: 'Лист с результатами не найден.' };
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  const backgrounds = range.getBackgrounds();

  if (data.length < 2) {
    return { error: 'Таблица пуста.' };
  }

  const header = data[0].map(String);
  const headerColors = backgrounds[0];

  let authCols;
  let allowedCols;

  try {
    authCols = getAuthColumnIndexes(header);
    allowedCols = getAllowedColumnIndexes(headerColors);
  } catch (_error) {
    logAccess({ login, password, clientInfo, status: 'Ошибка конфигурации столбцов' });
    return { error: 'Ошибка структуры таблицы.' };
  }

  const normalizedLogin = normalizeLogin(login);
  const snilsCol = getSnilsColumnIndex(header);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowLogin = normalizeLogin(row[authCols.loginCol]);
    const rowPassword = formatCellValue(row[authCols.passCol]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = snilsCol >= 0 ? normalizeSnils(row[snilsCol]) : '';
      if (rowSnils) {
        logAccess({ login, password, clientInfo, status: 'Требуется ввод СНИЛС' });
        return { requiresSnils: true };
      }

      logAccess({ login, password, clientInfo, status: 'Удачный вход' });
      return prepareRowForClient(row, header, backgrounds[i], allowedCols);
    }
  }

  logAccess({ login, password, clientInfo, status: 'Неудачный вход: ФИО/дата' });
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
  const sheet = getSheet(CONFIG.RESULT_SHEET_NAME);
  if (!sheet) {
    return { error: 'Лист с результатами не найден.' };
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  const backgrounds = range.getBackgrounds();

  if (data.length < 2) {
    return { error: 'Таблица пуста.' };
  }

  const header = data[0].map(String);
  const headerColors = backgrounds[0];

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

  const normalizedLogin = normalizeLogin(login);
  const expectedSnils = normalizeSnils(snils);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowLogin = normalizeLogin(row[authCols.loginCol]);
    const rowPassword = formatCellValue(row[authCols.passCol]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      const rowSnils = normalizeSnils(row[snilsCol]);
      if (!rowSnils) {
        logAccess({ login, password, clientInfo, status: 'Удачный вход без СНИЛС' });
        return prepareRowForClient(row, header, backgrounds[i], allowedCols);
      }

      if (rowSnils !== expectedSnils) {
        logAccess({ login, password, clientInfo, status: 'Неудачный вход: СНИЛС' });
        return { error: 'Неправильно введён СНИЛС' };
      }

      logAccess({ login, password, clientInfo, status: 'Удачный вход по СНИЛС' });
      return prepareRowForClient(row, header, backgrounds[i], allowedCols);
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

  return {
    selected,
    options: top.map(m => m.original)
  };
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
