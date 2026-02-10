/*************************************************
 * КОНСТАНТЫ
 *************************************************/
const SPREADSHEET_ID = '1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ';
const RESULT_SHEET_NAME = 'Результат';
const LOG_SHEET_NAME = 'Входы';
const YELLOW = '#ffff00';

/*************************************************
 * ТОЧКА ВХОДА WEB-APP
 *************************************************/
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('ДБВv5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtmlFile(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/*************************************************
 * ЛОГИРОВАНИЕ
 *************************************************/
function logAccess({ login = '', password = '', clientInfo = {}, status }) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);

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

function logFormClick(login) {
  logAccess({ login, status: 'Нажал: Заполнить форму' });
}

/*************************************************
 * ДОКУМЕНТЫ: АУТЕНТИФИКАЦИЯ/ДАННЫЕ
 *************************************************/
function getAllowedColumnIndexes(headerColors) {
  const result = [];
  for (let i = 0; i < headerColors.length; i++) {
    if (headerColors[i] === YELLOW) result.push(i);
  }
  return result;
}

function getAuthColumnIndexes(header) {
  const loginCol = header.indexOf('Фамилия Имя Отчество (С)');
  const passCol = header.indexOf('Дата рождения (С)');

  if (loginCol === -1 || passCol === -1) {
    throw new Error('AUTH_COLUMNS_NOT_FOUND');
  }

  return { loginCol, passCol };
}

function normalizeLogin(value) {
  return String(value || '').trim().toLowerCase().replace(/ё/g, 'е');
}

function formatCellValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'GMT+3', 'dd.MM.yyyy');
  }
  return String(value ?? '');
}

function prepareRowForClient(row, header, backgrounds, allowedCols) {
  return {
    header: allowedCols.map(i => header[i]),
    row: allowedCols.map(i => formatCellValue(row[i])),
    colors: allowedCols.map(i => backgrounds[i])
  };
}

function checkLogin(login, password, clientInfo = {}) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESULT_SHEET_NAME);

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
  } catch (e) {
    logAccess({ login, password, clientInfo, status: 'Ошибка конфигурации столбцов' });
    return { error: 'Ошибка структуры таблицы.' };
  }

  const normalizedLogin = normalizeLogin(login);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowLogin = normalizeLogin(row[authCols.loginCol]);
    const rowPassword = formatCellValue(row[authCols.passCol]).trim();

    if (rowLogin === normalizedLogin && rowPassword === password) {
      logAccess({ login, password, clientInfo, status: 'Удачный вход' });
      return prepareRowForClient(row, header, backgrounds[i], allowedCols);
    }
  }

  logAccess({ login, password, clientInfo, status: 'Неудачный вход' });
  return { error: 'Неверный логин или дата рождения.' };
}

/*************************************************
 * ПОСЕЩАЕМОСТЬ: ПОДБОР ФИО
 *************************************************/
function findNames(inputs) {
  const fullNames = loadFullNames();
  return inputs.map(input => processInput(input, fullNames));
}

function loadFullNames() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESULT_SHEET_NAME);

  return sheet
    .getRange(1, 1, sheet.getLastRow(), 1)
    .getValues()
    .flat()
    .filter(String)
    .map(normalizeFullName);
}

function processInput(input, fullNames) {
  const short = normalizeShortName(input);
  if (!short) return null;

  const maxErrors = 2;
  const matches = fullNames
    .map(full => calculateMatch(short, full, maxErrors))
    .filter(Boolean);

  if (!matches.length) return null;

  matches.sort((a, b) => a.totalCost - b.totalCost);

  const exactLast = matches.filter(m => m.lastCost === 0);
  const selected = exactLast.length === 1 ? exactLast[0].original : matches[0].original;

  const top = matches.slice(0, 3);
  if (!top.some(m => m.original === selected)) {
    const sel = matches.find(m => m.original === selected);
    if (sel) {
      top.pop();
      top.unshift(sel);
    }
  }

  return {
    selected,
    options: top.map(m => m.original)
  };
}

function normalize(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\./g, '')
    .replace(/\s+/g, '');
}

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

function normalizeShortName(text) {
  if (!text) return null;

  const clean = normalize(text);
  const m = clean.match(/^([а-я]+)([а-я]*)$/);
  if (!m) return null;

  return { last: m[1], tail: m[2] };
}

function fuzzyPrefixCost(text, pattern, maxErrors) {
  if (!pattern) return 0;
  if (!text) return Infinity;

  let min = Infinity;
  const minLen = Math.max(1, pattern.length - maxErrors);
  const maxLen = Math.min(text.length, pattern.length + maxErrors);

  for (let len = minLen; len <= maxLen; len++) {
    const part = text.slice(0, len);
    const d = levenshtein(part, pattern);
    if (d < min) min = d;
  }

  return min;
}

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
