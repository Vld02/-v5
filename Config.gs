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
  // Максимальная ширина сайта, px. При большей ширине экран сайт дальше не расширяется.
  pageMaxWidth: 1088,
  // Граница компактного режима в реальных CSS-пикселях viewport.
  // При большей ширине масштаб интерфейса всегда равен 100%. НЕ РАБОТАЕТ
  compactModeWidth: 1080,
  // Ширина viewport, при которой достигается минимальный масштаб интерфейса, px. НЕ РАБОТАЕТ
  compactMinViewportWidth: 1020,
  // Минимальный допустимый масштаб интерфейса (1 = 100%).
  minimumInterfaceScale: 1.00
});

// Журнал:
const LOG_CONFIG = Object.freeze({
  // Столбцы журнала:
  COLUMNS: Object.freeze([
    'Дата/время', 'Логин', 'Пароль', 'СНИЛС', 'IP', 'Устройство', 'Браузер',
    'Действие пользователя', 'Результат действия', 'Локальные данные'
  ]),
  EVENTS: Object.freeze({
    // Общие действия сайта:
    page_open: Object.freeze({ type: 'action', template: 'Открыл сайт' }),
    section_docs: Object.freeze({ type: 'action', template: 'Перешёл в раздел — Документы' }),
    section_attendance: Object.freeze({ type: 'action', template: 'Перешёл в раздел — Посещаемость' }),
    section_gear: Object.freeze({ type: 'action', template: 'Перешёл в раздел — Снаряжение' }),

    // Авторизация:
    login_click: Object.freeze({ type: 'action', template: 'Нажал: Войти' }),
    login_success_without_snils: Object.freeze({ type: 'result', template: 'Удачный вход без СНИЛС' }),
    login_success_with_snils: Object.freeze({ type: 'result', template: 'Удачный вход по СНИЛС' }),
    login_snils_required: Object.freeze({ type: 'result', template: 'Требуется ввод СНИЛС' }),
    login_invalid_snils: Object.freeze({ type: 'result', template: 'Неверный СНИЛС' }),
    login_failed_credentials: Object.freeze({ type: 'result', template: 'Неудачный вход: ФИО/дата' }),
    login_sheet_missing: Object.freeze({ type: 'result', template: 'Лист не найден' }),
    login_config_error: Object.freeze({ type: 'result', template: 'Ошибка конфигурации столбцов' }),
    user_warning: Object.freeze({ type: 'result', template: 'Предупреждение: {message}' }),
    user_error: Object.freeze({ type: 'result', template: 'Ошибка: {message}' }),

    // Документы: редактирование:
    edit_click: Object.freeze({ type: 'action', template: 'Нажал: Редактировать — {field}' }),
    save_click: Object.freeze({ type: 'action', template: 'Нажал: Сохранить — {field}' }),
    cancel_click: Object.freeze({ type: 'action', template: 'Нажал: Отмена — {field}' }),

    // Документы: локальные данные:
    draft_saved: Object.freeze({ type: 'local', template: 'Локально сохранено: {field} → {value}' }),

    // Документы: серверное сохранение:
    save_success: Object.freeze({ type: 'result', template: 'Сохранено на сервере: {field} → {value}' }),
    save_error: Object.freeze({ type: 'result', template: 'Ошибка сохранения: {field}' }),

    // Документы: файлы:
    file_open: Object.freeze({ type: 'action', template: 'Открыл файл — {field}' }),
    file_open_error: Object.freeze({ type: 'result', template: 'Ошибка открытия файла — {field}' }),
    file_attach_click: Object.freeze({ type: 'action', template: 'Нажал: Прикрепить файл — {field}' }),
    file_selected: Object.freeze({ type: 'action', template: 'Выбрал файл — {fileName}' }),
    file_upload_success: Object.freeze({ type: 'result', template: 'Файл загружен: {field}' }),
    file_upload_error: Object.freeze({ type: 'result', template: 'Ошибка загрузки файла: {field}' }),
    file_replace_continue: Object.freeze({ type: 'action', template: 'Продолжил замену файла — {field}' }),
    file_replace_cancel: Object.freeze({ type: 'action', template: 'Отменил замену файла — {field}' }),

    // Посещаемость: работа с ячейками:
    attendance_cell_change: Object.freeze({ type: 'action', template: 'Заполнил ячейку — {value}' }),

    // Посещаемость: выбор ФИО:
    fio_select: Object.freeze({ type: 'action', template: 'Выбрал ФИО — {fio}' }),
    fio_change: Object.freeze({ type: 'action', template: 'Изменил выбор ФИО — {oldFio} → {fio}' }),

    // Посещаемость: буфер обмена:
    attendance_paste_click: Object.freeze({ type: 'action', template: 'Нажал: Вставить из буфера' }),
    attendance_paste_error: Object.freeze({ type: 'result', template: 'Ошибка вставки из буфера' }),
    attendance_copy_click: Object.freeze({ type: 'action', template: 'Нажал: Скопировать полное ФИО' }),
    attendance_copy_success: Object.freeze({ type: 'result', template: 'Полные ФИО скопированы' }),
    attendance_copy_error: Object.freeze({ type: 'result', template: 'Ошибка копирования ФИО' }),

    // Посещаемость: заполнение тренировки:
    attendance_form_click: Object.freeze({ type: 'action', template: 'Нажал: Заполнить тренировку' }),
    attendance_form_open: Object.freeze({ type: 'result', template: 'Открыл форму заполнения тренировки' })
  }),
  // Ожидание записи журнала, миллисекунд:
  lockWaitMs: 1000
});
const LOG_COLUMNS = LOG_CONFIG.COLUMNS;

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
