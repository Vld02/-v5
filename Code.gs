/*************************************************
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 *************************************************/
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
