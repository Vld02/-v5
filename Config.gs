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
    'Дата/время', 'Логин', 'Пароль', 'СНИЛС', 'IP', 'Устройство', 'Браузер', 'Действие пользователя', 'Результат действия', 'Локальные данные'
  ]),
  // Время объединения записей, минут:
  MAX_AGE_MINUTES: 30,
  // Ожидание записи журнала, миллисекунд:
  lockWaitMs: 1000
});
const LOG_COLUMNS = LOG_CONFIG.COLUMNS;
const LOG_MAX_AGE_MINUTES = LOG_CONFIG.MAX_AGE_MINUTES;

// Технические настройки внутренней связи логического события со строкой
// журнала. Метаданные диапазона не образуют видимого столбца и перемещаются
// вместе со строкой при её вставке или перемещении в таблице.
const LOGICAL_EVENT_LOG_CONFIG = Object.freeze({
  eventIdMetadataKey: 'logical-log-event-id',
  sessionIdMetadataKey: 'logical-log-session-id',
  stateMetadataKey: 'logical-log-event-state',
  completedState: 'completed'
});

// Конфигурация логических событий пользовательского журнала.
//
// Это единственный источник текстов и правил пользовательских событий.
// Идентификатор события всегда создаётся отдельно от номера строки листа.
const LOG_EVENT_CONFIG = Object.freeze({
  EVENT_TYPES: Object.freeze({
    ACTION: 'action',
    RESULT: 'result',
    LOCAL_DATA: 'local_data'
  }),
  SOURCES: Object.freeze({
    CLIENT: 'index.html',
    SERVER: 'Code.gs'
  }),
  RULES: Object.freeze({
    CREATE_NEW: 'create_new_event',
    ATTACH_TO_EVENT: 'attach_to_event'
  }),
  EVENTS: Object.freeze({
    // Общие события.
    page_open: Object.freeze({ id: 'page_open', name: 'Открыл сайт', type: 'action', source: 'index.html', template: 'Открыл сайт', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    section_docs: Object.freeze({ id: 'section_docs', name: 'Переход в раздел «Документы»', type: 'action', source: 'index.html', template: 'Перешёл в раздел — Документы', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    section_attendance: Object.freeze({ id: 'section_attendance', name: 'Переход в раздел «Посещаемость»', type: 'action', source: 'index.html', template: 'Перешёл в раздел — Посещаемость', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    section_gear: Object.freeze({ id: 'section_gear', name: 'Переход в раздел «Снаряжение»', type: 'action', source: 'index.html', template: 'Перешёл в раздел — Снаряжение', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),

    // Авторизация. Успешными вариантами являются только два события ниже.
    login_click: Object.freeze({ id: 'login_click', name: 'Нажатие кнопки «Войти»', type: 'action', source: 'index.html', template: 'Нажал: Войти', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    login_success_without_snils: Object.freeze({ id: 'login_success_without_snils', name: 'Удачный вход без СНИЛС', type: 'result', source: 'Code.gs', template: 'Удачный вход без СНИЛС', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_success_with_snils: Object.freeze({ id: 'login_success_with_snils', name: 'Удачный вход по СНИЛС', type: 'result', source: 'Code.gs', template: 'Удачный вход по СНИЛС', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_snils_required: Object.freeze({ id: 'login_snils_required', name: 'Требуется ввод СНИЛС', type: 'result', source: 'Code.gs', template: 'Требуется ввод СНИЛС', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_invalid_snils: Object.freeze({ id: 'login_invalid_snils', name: 'Неверный СНИЛС', type: 'result', source: 'Code.gs', template: 'Неверный СНИЛС', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_failed_credentials: Object.freeze({ id: 'login_failed_credentials', name: 'Неверные ФИО или дата рождения', type: 'result', source: 'Code.gs', template: 'Неудачный вход: ФИО/дата', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_sheet_missing: Object.freeze({ id: 'login_sheet_missing', name: 'Лист авторизации не найден', type: 'result', source: 'Code.gs', template: 'Лист не найден', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),
    login_config_error: Object.freeze({ id: 'login_config_error', name: 'Ошибка конфигурации авторизации', type: 'result', source: 'Code.gs', template: 'Ошибка конфигурации столбцов', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'login_click', enabled: true, completeAfterWrite: true}),

    // Пользовательские сообщения.
    user_warning: Object.freeze({ id: 'user_warning', name: 'Предупреждение пользователю', type: 'result', source: 'index.html', template: 'Предупреждение: {message}', params: Object.freeze(['message']), mode: 'attach_to_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    user_error: Object.freeze({ id: 'user_error', name: 'Ошибка пользователю', type: 'result', source: 'index.html', template: 'Ошибка: {message}', params: Object.freeze(['message']), mode: 'attach_to_event', parentEvent: null, enabled: true, completeAfterWrite: true}),

    // Документы.
    edit_click: Object.freeze({ id: 'edit_click', name: 'Нажатие «Редактировать»', type: 'action', source: 'index.html', template: 'Нажал: Редактировать — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    save_click: Object.freeze({ id: 'save_click', name: 'Нажатие «Сохранить»', type: 'action', source: 'index.html', template: 'Нажал: Сохранить — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    cancel_click: Object.freeze({ id: 'cancel_click', name: 'Нажатие «Отмена»', type: 'action', source: 'index.html', template: 'Нажал: Отмена — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    draft_saved: Object.freeze({ id: 'draft_saved', name: 'Локальное сохранение редактора', type: 'local_data', source: 'index.html', template: 'Локально сохранено: {field} → {value}', params: Object.freeze(['field', 'value']), mode: 'attach_to_event', parentEvent: 'edit_click', enabled: true, completeAfterWrite: true}),
    save_success: Object.freeze({ id: 'save_success', name: 'Документ сохранён на сервере', type: 'result', source: 'Code.gs', template: 'Сохранено на сервере: {field} → {value}', params: Object.freeze(['field', 'value']), mode: 'attach_to_event', parentEvent: 'save_click', enabled: true, completeAfterWrite: true}),
    save_error: Object.freeze({ id: 'save_error', name: 'Ошибка сохранения документа', type: 'result', source: 'Code.gs', template: 'Ошибка сохранения: {field}', params: Object.freeze(['field']), mode: 'attach_to_event', parentEvent: 'save_click', enabled: true, completeAfterWrite: true}),

    // Файлы.
    file_open: Object.freeze({ id: 'file_open', name: 'Открытие файла', type: 'action', source: 'index.html', template: 'Открыл файл — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    file_open_error: Object.freeze({ id: 'file_open_error', name: 'Ошибка открытия файла', type: 'result', source: 'index.html', template: 'Ошибка открытия файла — {field}', params: Object.freeze(['field']), mode: 'attach_to_event', parentEvent: 'file_open', enabled: true, completeAfterWrite: true}),
    file_attach_click: Object.freeze({ id: 'file_attach_click', name: 'Нажатие «Прикрепить файл»', type: 'action', source: 'index.html', template: 'Нажал: Прикрепить файл — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    file_selected: Object.freeze({ id: 'file_selected', name: 'Выбор файла', type: 'action', source: 'index.html', template: 'Выбрал файл — {fileName}', params: Object.freeze(['fileName']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    file_upload_success: Object.freeze({ id: 'file_upload_success', name: 'Файл загружен', type: 'result', source: 'Code.gs', template: 'Файл загружен: {field}', params: Object.freeze(['field']), mode: 'attach_to_event', parentEvent: 'file_selected', enabled: true, completeAfterWrite: true}),
    file_upload_error: Object.freeze({ id: 'file_upload_error', name: 'Ошибка загрузки файла', type: 'result', source: 'Code.gs', template: 'Ошибка загрузки файла: {field}', params: Object.freeze(['field']), mode: 'attach_to_event', parentEvent: 'file_selected', enabled: true, completeAfterWrite: true}),
    file_replace_continue: Object.freeze({ id: 'file_replace_continue', name: 'Продолжение замены файла', type: 'action', source: 'index.html', template: 'Продолжил замену файла — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    file_replace_cancel: Object.freeze({ id: 'file_replace_cancel', name: 'Отмена замены файла', type: 'action', source: 'index.html', template: 'Отменил замену файла — {field}', params: Object.freeze(['field']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),

    // Посещаемость.
    attendance_cell_change: Object.freeze({ id: 'attendance_cell_change', name: 'Заполнение ячейки посещаемости', type: 'action', source: 'index.html', template: 'Заполнил ячейку — {value}', params: Object.freeze(['value']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    fio_select: Object.freeze({ id: 'fio_select', name: 'Выбор полного ФИО', type: 'action', source: 'index.html', template: 'Выбрал ФИО — {fio}', params: Object.freeze(['fio']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    fio_change: Object.freeze({ id: 'fio_change', name: 'Изменение выбранного полного ФИО', type: 'action', source: 'index.html', template: 'Изменил выбор ФИО — {previousFio} → {fio}', params: Object.freeze(['previousFio', 'fio']), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: true}),
    attendance_paste_click: Object.freeze({ id: 'attendance_paste_click', name: 'Нажатие «Вставить из буфера»', type: 'action', source: 'index.html', template: 'Нажал: Вставить из буфера', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    attendance_paste_error: Object.freeze({ id: 'attendance_paste_error', name: 'Ошибка вставки из буфера', type: 'result', source: 'index.html', template: 'Ошибка вставки из буфера', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'attendance_paste_click', enabled: true, completeAfterWrite: true}),
    attendance_copy_click: Object.freeze({ id: 'attendance_copy_click', name: 'Нажатие «Скопировать полное ФИО»', type: 'action', source: 'index.html', template: 'Нажал: Скопировать полное ФИО', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    attendance_copy_success: Object.freeze({ id: 'attendance_copy_success', name: 'Полные ФИО скопированы', type: 'result', source: 'index.html', template: 'Полные ФИО скопированы', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'attendance_copy_click', enabled: true, completeAfterWrite: true}),
    attendance_copy_error: Object.freeze({ id: 'attendance_copy_error', name: 'Ошибка копирования ФИО', type: 'result', source: 'index.html', template: 'Ошибка копирования ФИО', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'attendance_copy_click', enabled: true, completeAfterWrite: true}),
    attendance_form_click: Object.freeze({ id: 'attendance_form_click', name: 'Нажатие «Заполнить тренировку»', type: 'action', source: 'index.html', template: 'Нажал: Заполнить тренировку', params: Object.freeze([]), mode: 'create_new_event', parentEvent: null, enabled: true, completeAfterWrite: false}),
    attendance_form_open: Object.freeze({ id: 'attendance_form_open', name: 'Открытие формы заполнения тренировки', type: 'result', source: 'index.html', template: 'Открыл форму заполнения тренировки', params: Object.freeze([]), mode: 'attach_to_event', parentEvent: 'attendance_form_click', enabled: true, completeAfterWrite: true})
  })
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
