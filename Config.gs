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
    'Дата/время', 'Логин', 'Пароль', 'СНИЛС', 'IP',
    'Устройство', 'Браузер', 'Действие пользователя',
    'Результат действия', 'Локальные данные'
  ]),
  // Время объединения записей, минут:
  MAX_AGE_MINUTES: 30,
  // Ожидание записи журнала, миллисекунд:
  lockWaitMs: 1000
});
const LOG_COLUMNS = LOG_CONFIG.COLUMNS;
const LOG_MAX_AGE_MINUTES = LOG_CONFIG.MAX_AGE_MINUTES;

/*************************************************
 * КОНФИГУРАЦИЯ ЛОГИЧЕСКИХ СОБЫТИЙ ЖУРНАЛА «ВХОДЫ»
 *
 * Эта конфигурация подготовлена для нового механизма журналирования.
 * На этапе подготовки она намеренно не используется существующими
 * функциями logPageOpen(), logSessionEvent() и logAuthAttempt().
 * Поэтому текущая схема записи журнала остаётся без изменений.
 *************************************************/
const USER_LOG_EVENT_TYPES = Object.freeze({
  ACTION: 'action',
  RESULT: 'result',
  LOCAL_DATA: 'local_data'
});

const USER_LOG_EVENT_SOURCES = Object.freeze({
  CLIENT: 'index.html',
  SERVER: 'Code.gs'
});

const USER_LOG_EVENT_RULES = Object.freeze({
  CREATE: 'create_new_event',
  JOIN: 'join_current_event'
});

/**
 * Единая видимая структура пользовательского журнала без столбца внутреннего ID.
 */
const USER_LOG_COLUMNS = LOG_CONFIG.COLUMNS;

/**
 * Описание событий нового пользовательского журнала.
 *
 * parameters — допустимые подстановки шаблона {parameter}.
 * joinToEventTypes — типы текущего события, к которым разрешено присоединить
 * результат или локальные данные. Пустой список означает создание новой строки.
 */
const USER_LOG_EVENTS = Object.freeze({
  page_open: Object.freeze({ id: 'page_open', name: 'Открытие страницы', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Открыл страницу', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  section_docs: Object.freeze({ id: 'section_docs', name: 'Переход в раздел «Документы»', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Перешёл в раздел «Документы»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  section_attendance: Object.freeze({ id: 'section_attendance', name: 'Переход в раздел «Посещаемость»', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Перешёл в раздел «Посещаемость»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  section_gear: Object.freeze({ id: 'section_gear', name: 'Переход в раздел «Снаряжение»', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Перешёл в раздел «Снаряжение»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),

  login_click: Object.freeze({ id: 'login_click', name: 'Нажатие «Войти»', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Войти»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  login_success_without_snils: Object.freeze({ id: 'login_success_without_snils', name: 'Успешный вход без СНИЛС', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Вход выполнен без СНИЛС', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_success_with_snils: Object.freeze({ id: 'login_success_with_snils', name: 'Успешный вход со СНИЛС', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Вход выполнен со СНИЛС', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_snils_required: Object.freeze({ id: 'login_snils_required', name: 'Требуется СНИЛС', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Требуется СНИЛС', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_invalid_snils: Object.freeze({ id: 'login_invalid_snils', name: 'Неверный СНИЛС', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Неверный СНИЛС', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_failed_credentials: Object.freeze({ id: 'login_failed_credentials', name: 'Неверные учётные данные', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Неверные ФИО или дата рождения', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_sheet_missing: Object.freeze({ id: 'login_sheet_missing', name: 'Лист авторизации не найден', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Лист авторизации не найден', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  login_config_error: Object.freeze({ id: 'login_config_error', name: 'Ошибка конфигурации авторизации', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Ошибка конфигурации авторизации: {message}', parameters: Object.freeze(['message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),

  user_warning: Object.freeze({ id: 'user_warning', name: 'Предупреждение пользователю', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Предупреждение: {message}', parameters: Object.freeze(['message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  user_error: Object.freeze({ id: 'user_error', name: 'Ошибка для пользователя', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Ошибка: {message}', parameters: Object.freeze(['message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),

  edit_click: Object.freeze({ id: 'edit_click', name: 'Начало редактирования документа', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Редактировать»: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  save_click: Object.freeze({ id: 'save_click', name: 'Сохранение документа', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Сохранить»: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  cancel_click: Object.freeze({ id: 'cancel_click', name: 'Отмена редактирования', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Отмена»: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  draft_saved: Object.freeze({ id: 'draft_saved', name: 'Черновик сохранён локально', type: USER_LOG_EVENT_TYPES.LOCAL_DATA, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Черновик сохранён: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  save_success: Object.freeze({ id: 'save_success', name: 'Документ сохранён', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Сохранено: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  save_error: Object.freeze({ id: 'save_error', name: 'Ошибка сохранения документа', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Не сохранено: {field}; {message}', parameters: Object.freeze(['field', 'message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),

  file_open: Object.freeze({ id: 'file_open', name: 'Открытие файла', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Открыл файл: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  file_open_error: Object.freeze({ id: 'file_open_error', name: 'Ошибка открытия файла', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Не открыт файл: {field}; {message}', parameters: Object.freeze(['field', 'message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  file_attach_click: Object.freeze({ id: 'file_attach_click', name: 'Нажатие прикрепления файла', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Прикрепить файл»: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  file_selected: Object.freeze({ id: 'file_selected', name: 'Файл выбран локально', type: USER_LOG_EVENT_TYPES.LOCAL_DATA, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Выбран файл: {fileName}', parameters: Object.freeze(['fileName']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  file_upload_success: Object.freeze({ id: 'file_upload_success', name: 'Файл загружен', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Файл загружен: {fileName}', parameters: Object.freeze(['fileName']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  file_upload_error: Object.freeze({ id: 'file_upload_error', name: 'Ошибка загрузки файла', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.SERVER, template: 'Файл не загружен: {fileName}; {message}', parameters: Object.freeze(['fileName', 'message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  file_replace_continue: Object.freeze({ id: 'file_replace_continue', name: 'Подтверждение замены файла', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Продолжил замену файла: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  file_replace_cancel: Object.freeze({ id: 'file_replace_cancel', name: 'Отмена замены файла', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Отменил замену файла: {field}', parameters: Object.freeze(['field']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),

  attendance_cell_change: Object.freeze({ id: 'attendance_cell_change', name: 'Изменение ячейки посещаемости', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Изменил посещаемость: {date}, {athlete}', parameters: Object.freeze(['date', 'athlete']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  fio_select: Object.freeze({ id: 'fio_select', name: 'Выбор ФИО', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Выбрал ФИО: {fio}', parameters: Object.freeze(['fio']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  fio_change: Object.freeze({ id: 'fio_change', name: 'Изменение ФИО', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Изменил ФИО: {fio}', parameters: Object.freeze(['fio']), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  attendance_paste_click: Object.freeze({ id: 'attendance_paste_click', name: 'Вставка посещаемости', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Вставить посещаемость»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  attendance_paste_error: Object.freeze({ id: 'attendance_paste_error', name: 'Ошибка вставки посещаемости', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Не вставлена посещаемость: {message}', parameters: Object.freeze(['message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  attendance_copy_click: Object.freeze({ id: 'attendance_copy_click', name: 'Копирование посещаемости', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Копировать посещаемость»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  attendance_copy_success: Object.freeze({ id: 'attendance_copy_success', name: 'Посещаемость скопирована', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Посещаемость скопирована', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  attendance_copy_error: Object.freeze({ id: 'attendance_copy_error', name: 'Ошибка копирования посещаемости', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Не скопирована посещаемость: {message}', parameters: Object.freeze(['message']), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true }),
  attendance_form_click: Object.freeze({ id: 'attendance_form_click', name: 'Нажатие формы посещаемости', type: USER_LOG_EVENT_TYPES.ACTION, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Нажал «Заполнить форму»', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.CREATE, joinToEventTypes: Object.freeze([]), enabled: true }),
  attendance_form_open: Object.freeze({ id: 'attendance_form_open', name: 'Форма посещаемости открыта', type: USER_LOG_EVENT_TYPES.RESULT, source: USER_LOG_EVENT_SOURCES.CLIENT, template: 'Форма посещаемости открыта', parameters: Object.freeze([]), rule: USER_LOG_EVENT_RULES.JOIN, joinToEventTypes: Object.freeze([USER_LOG_EVENT_TYPES.ACTION]), enabled: true })
});

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
