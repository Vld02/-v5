/*************************************************
 * ОБЩИЕ НАСТРОЙКИ САЙТА
 *
 * Меняйте значения в конфигурации, не меняя названия свойств.
 *************************************************/

// Ссылка на иконку сайта:
const WEBAPP_FAVICON_URL = 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.ico';

// Основные настройки сайта:
const CONFIG = Object.freeze({
  // Ссылка на основную таблицу Google Sheets:
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ/edit',
  // Название листа с персональными данными:
  RESULT_SHEET_NAME: 'Результат',
  // Название листа журнала входов на сайт:
  LOG_SHEET_NAME: 'Входы',
  // Цвет показываемого столбца:
  YELLOW: '#ffff00',
  // Часовой пояс приложения:
  TIMEZONE: 'GMT+3',
  // Формат даты:
  DATE_FORMAT: 'dd.MM.yyyy',
  // Общая авторизация:
  AUTH: Object.freeze({
    // Столбец с ФИО для входа:
    loginHeader: 'Фамилия Имя Отчество (С)',
    // Столбец с датой рождения для входа:
    passwordHeader: 'Дата рождения (С)',
    // Столбец со СНИЛС:
    snilsHeader: 'Снилс: номер (С)'
  })
});

// Настройки приложения:
const APP_CONFIG = Object.freeze({
  // Название сайта:
  APP_TITLE: 'ДБВv5',
  // Цвет темы сайта:
  THEME_COLOR: '#179bcf',
  // Ссылки на иконки сайта:
  ICON_URLS: Object.freeze({
    // Иконка 32×32:
    icon32: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/32.png',
    // Иконка 72×72:
    icon72: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/72.png',
    // Иконка 192×192:
    icon192: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/192.png',
    // Иконка 512×512:
    icon512: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.png'
  }),
  // Названия разделов сайта:
  SECTION_NAMES: Object.freeze({
    // Раздел документов:
    docs: 'ДБВv5 Документы',
    // Раздел посещаемости:
    attendance: 'ДБВv5 Посещаемость',
    // Раздел снаряжения:
    gear: 'ДБВv5 Снаряжение'
  })
});

// Адаптивность основного блока сайта:
const RESPONSIVE_CONFIG = Object.freeze({
  // Максимальная ширина сайта, px. При большей ширине экрана сайт дальше не расширяется.
  pageMaxWidth: 1088,
  // Граница компактного режима, px. Ниже этой ширины начинает уменьшаться масштаб интерфейса.
  compactModeWidth: 100
});

// Журнал:
const LOG_CONFIG = Object.freeze({
  // Столбцы журнала:
  COLUMNS: Object.freeze([
    'Дата/время входа', 'Логин', 'Пароль', 'СНИЛС', 'IP', 'Устройство', 'Браузер', 'Статус входа'
  ]),
  // Время объединения записей, минут:
  MAX_AGE_MINUTES: 30,
  // Ожидание записи журнала, миллисекунд:
  lockWaitMs: 1000
});
const LOG_COLUMNS = LOG_CONFIG.COLUMNS;
const LOG_MAX_AGE_MINUTES = LOG_CONFIG.MAX_AGE_MINUTES;

// Общие настройки интерфейса браузера:
const GENERAL_CLIENT_CONFIG = Object.freeze({
  // Уникальные ключи для сохранения данных сайта в браузере:
  storageKeys: Object.freeze({
    savedLogin: 'savedLogin',
    savedDate: 'savedDate',
    savedSnils: 'savedSnils',
    isLoggedIn: 'isLoggedIn',
    activeTab: 'activeTab'
  }),
  // Общие интервалы интерфейса, миллисекунд:
  timings: Object.freeze({
    // Через сколько скрывать обычное уведомление:
    toastAutoHideMs: 3500,
    // Скорость анимации загрузки при входе:
    authLoadingAnimationMs: 450,
    // Пауза перед закрытием списка подсказок, миллисекунд:
    suggestionBlurDelayMs: 120
  }),
  // Общие ссылки сайта:
  urls: Object.freeze({
    // Сервис определения IP-адреса пользователя:
    ipLookup: 'https://api.ipify.org?format=json'
  })
});
