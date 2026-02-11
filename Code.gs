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

/*************************************************
 * СИНХРОНИЗАЦИЯ GOOGLE FORM ИЗ ЛИСТА "Списки данных"
 *************************************************/
// @ts-ignore
function syncFormFromSheet() {
  const FORM_ID = '1_em0kZ8lzw1blqKpDWslBlowb-lXtqxovOfSuZiRqhk';
  const SHEET_ID = '1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ';
  const DATA_SHEET = 'Списки данных';

  const form = FormApp.openById(FORM_ID);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET);

  if (!sheet) {
    Logger.log('❌ Лист "Списки данных" не найден');
    return;
  }

  Logger.log('▶ Старт синхронизации формы');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log('📋 Заголовки таблицы:');
  headers.forEach((h, i) => Logger.log(`  [${i}] "${h}"`));

  const getCol = name => headers.indexOf(name);

  function getColumnValues(name) {
    const idx = getCol(name);
    if (idx === -1) {
      Logger.log(`❌ Столбец не найден: "${name}"`);
      return null;
    }
    const values = data.slice(1).map(r => r[idx]).filter(v => v && v !== '×');
    if (!values.length) {
      Logger.log(`⚠ Столбец "${name}" найден, но данных нет`);
    }
    return values;
  }

  const items = form.getItems();
  const sections = items.filter(i => i.getType() === FormApp.ItemType.PAGE_BREAK);

  /* =======================================================
     1. Тренер присутствовал:
  ======================================================= */

  const trainerQ = items.find(i => i.getTitle() === 'Тренер присутствовал:');

  if (trainerQ && trainerQ.getType() === FormApp.ItemType.CHECKBOX) {
    const values = sheet
      .getRange('F2:F')
      .getValues()
      .flat()
      .filter(v => v && v !== '×');

    if (values.length) {
      const q = trainerQ.asCheckboxItem();
      q.setChoices(values.map(v => q.createChoice(v)));
      Logger.log(`✅ "Тренер присутствовал:" (${values.length})`);
    } else {
      Logger.log('⚠ Диапазон F2:F не содержит данных для "Тренер присутствовал:"');
    }
  } else {
    Logger.log('⚠ Вопрос "Тренер присутствовал:" не найден или имеет неверный тип');
  }

  /* =======================================================
     2. Группы по расписанию
  ======================================================= */

  const groupQ = items.find(i =>
    i.getTitle() === 'Группы у которых проходила тренировка по расписанию'
  );

  if (groupQ && groupQ.getType() === FormApp.ItemType.CHECKBOX) {
    const values = getColumnValues('Группы тренировки');
    if (values?.length) {
      const q = groupQ.asCheckboxItem();
      q.setChoices(values.map(v => q.createChoice(v)));
      Logger.log(`✅ "Группы по расписанию" (${values.length})`);
    }
  }

  /* =======================================================
   3–4. Разделы 5–14 (один вопрос, множественный выбор)
======================================================= */

  const partColIdx = getCol('Части тренировок и группы');
  let partNames = [];

  if (partColIdx === -1) {
    Logger.log('❌ Столбец "Части тренировок и группы" не найден');
  } else {
    // строки 3–12 → названия вопросов 5–14
    partNames = data.slice(2, 12).map(r => r[partColIdx]).filter(Boolean);
    Logger.log(`ℹ Названия вопросов (5–14): ${partNames.length}`);
  }

  const sectionTitles = [
    '1 группа начальной подготовки',
    '2 группа начальной подготовки',
    '3 группа начальной подготовки',
    '4 группа начальной подготовки',
    '5 группа начальной подготовки',
    '1 тренировочная группа',
    '2 тренировочная группа',
    '3 тренировочная группа',
    '4 тренировочная группа',
    '5 тренировочная группа'
  ];

  sectionTitles.forEach((title, idx) => {
    const section = sections.find(s => s.getTitle() === title);
    if (!section) {
      Logger.log(`⚠ Раздел не найден: "${title}"`);
      return;
    }

    // берем первый вопрос после PAGE_BREAK
    const qItem = items[items.indexOf(section) + 1];
    if (!qItem) {
      Logger.log(`⚠ В разделе "${title}" нет вопроса`);
      return;
    }
    if (qItem.getType() !== FormApp.ItemType.CHECKBOX) {
      Logger.log(`⚠ Вопрос в разделе "${title}" не CHECKBOX`);
      return;
    }

    // берем варианты из колонки с названием раздела
    const colIdx = getCol(title);
    if (colIdx === -1) {
      Logger.log(`⚠ Колонка для раздела "${title}" не найдена`);
      return;
    }

    const values = data.slice(1).map(r => r[colIdx]).filter(v => v && v !== '×');
    if (!values.length) {
      Logger.log(`⚠ В разделе "${title}" нет данных для вариантов`);
      return;
    }

    const q = qItem.asCheckboxItem();

    // меняем название вопроса, если есть соответствующее значение из partNames
    if (partNames[idx]) q.setTitle(partNames[idx]);

    // выставляем варианты
    q.setChoices(values.map(v => q.createChoice(v)));

    Logger.log(`✅ Раздел "${title}" обновлен (${values.length} вариантов)`);
  });


  /* =======================================================
   Навигация для "Части тренировок и группы" (ИСПРАВЛЕНО)
======================================================= */

  const navQ = items.find(i => i.getTitle() === 'Части тренировок и группы');

  if (!navQ || navQ.getType() !== FormApp.ItemType.MULTIPLE_CHOICE) {
    Logger.log('❌ Навигационный вопрос не найден или имеет неверный тип');
  } else {

    const values = getColumnValues('Части тренировок и группы');
    if (!values || !values.length) {
      Logger.log('⏭ Навигация пропущена — нет данных в столбце');
    } else {

      const q = navQ.asMultipleChoiceItem();
      const choices = [];

      /* ===== 1️⃣ Первый вариант → "Программа и результаты" ===== */

      const section4 = sections.find(s => s.getTitle() === 'Программа и результаты');
      if (values[0] && section4) {
        choices.push(
          q.createChoice(values[0], section4.asPageBreakItem())
        );
      }

      /* ===== 2️⃣ СЕРЕДИНА → по совпадению с названием ВОПРОСА (разделы 5–14) ===== */

      const middle = values.slice(1, -3);

      middle.forEach(v => {
        if (!v) return;

        let targetSection = null;

        for (const section of sections) {
          const idx = items.indexOf(section);
          const nextItem = items[idx + 1];

          // первый вопрос после PAGE_BREAK
          if (
            nextItem &&
            nextItem.getType() === FormApp.ItemType.CHECKBOX &&
            nextItem.getTitle() === v
          ) {
            targetSection = section;
            break;
          }
        }

        if (targetSection) {
          choices.push(
            q.createChoice(v, targetSection.asPageBreakItem())
          );
        } else {
          Logger.log(`⚠ Не найден вопрос 5–14 с названием "${v}"`);
        }
      });

      /* ===== 3️⃣ Последние три варианта (В КОНЦЕ!) ===== */

      const last3 = values.slice(-3);

      const section15 = sections.find(s => s.getTitle() === 'Все группы');
      if (last3[0] && section15) {
        choices.push(
          q.createChoice(last3[0], section15.asPageBreakItem())
        );
      }

      const section2 = sections.find(s => s.getTitle() === 'Посещаемость');
      if (last3[1] && section2) {
        choices.push(
          q.createChoice(last3[1], section2.asPageBreakItem())
        );
      }

      if (last3[2]) {
        choices.push(
          q.createChoice(last3[2], FormApp.PageNavigationType.SUBMIT)
        );
      }

      q.setChoices(choices);
      Logger.log(`✅ Навигация создана корректно (${choices.length} вариантов)`);
    }
  }


  /* =======================================================
   7. Раздел "Все группы" — копия вопросов 5–14
======================================================= */

  const allSection = sections.find(s => s.getTitle() === 'Все группы');

  if (!allSection) {
    Logger.log('⚠ Раздел "Все группы" не найден');
  } else {

    /* === 1️⃣ Собираем эталонные вопросы из разделов 5–14 === */

    const sourceQuestions = [];

    sectionTitles.forEach(title => {
      const section = sections.find(s => s.getTitle() === title);
      if (!section) return;

      const idx = items.indexOf(section);
      const qItem = items[idx + 1];

      if (qItem && qItem.getType() === FormApp.ItemType.CHECKBOX) {
        const q = qItem.asCheckboxItem();
        sourceQuestions.push({
          title: q.getTitle(),
          choices: q.getChoices().map(c => c.getValue())
        });
      }
    });

    if (sourceQuestions.length !== 10) {
      Logger.log(`⚠ Ожидалось 10 эталонных вопросов, найдено ${sourceQuestions.length}`);
    }

    /* === 2️⃣ Вопросы в разделе "Все группы" === */

    const start = items.indexOf(allSection) + 1;
    const targetQuestions = items
      .slice(start)
      .filter(i => i.getType() === FormApp.ItemType.CHECKBOX)
      .slice(0, sourceQuestions.length);

    /* === 3️⃣ Копируем название и варианты === */

    targetQuestions.forEach((qItem, idx) => {
      const src = sourceQuestions[idx];
      if (!src) return;

      const q = qItem.asCheckboxItem();

      q.setTitle(src.title);
      q.setChoices(src.choices.map(v => q.createChoice(v)));

      Logger.log(`✅ Все группы: скопирован вопрос "${src.title}"`);
    });
  }
}
