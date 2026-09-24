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
  // Физическая структура листа «Входы».
  COLUMNS: Object.freeze([
    'Дата/время', 'Логин', 'Пароль', 'СНИЛС', 'IP', 'Устройство', 'Браузер',
    'Действие пользователя', 'Результат действия', 'Локальные данные'
  ]),

  // Соответствие типа события столбцу журнала.
  COLUMN_BY_TYPE: Object.freeze({
    action: 'Действие пользователя',
    result: 'Результат действия',
    local: 'Локальные данные'
  }),

  // Технические настройки.
  lockWaitMs: 1000,
  sessionNotePrefix: 'ДБВv5 session: ',

  // Каталог событий.
  // enabled      — разрешено ли событие;
  // type         — action / result / local;
  // template     — шаблон текста, поддерживает {field}, {value}, {section}, {text};
  // description  — когда событие должно возникать.
  EVENTS: Object.freeze({
    page_open_action: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Открыл сайт',
      description: 'Создание новой сессии при открытии страницы.'
    }),
    page_open_result: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Зашел на сайт',
      description: 'Результат первоначального открытия сайта.'
    }),

    login_click: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Войти',
      description: 'Нажатие кнопки входа.'
    }),
    auth_success: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Удачный вход',
      description: 'Успешная авторизация по ФИО и дате рождения.'
    }),
    auth_snils_required: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Требуется ввод СНИЛС',
      description: 'ФИО и дата совпали, требуется СНИЛС.'
    }),
    auth_invalid_snils: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Неверный СНИЛС',
      description: 'Введённый СНИЛС не совпал с данными.'
    }),
    auth_credentials_invalid: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Неудачный вход: ФИО/дата',
      description: 'ФИО и/или дата рождения не совпали.'
    }),
    auth_success_no_snils: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Удачный вход без СНИЛС',
      description: 'В строке пользователя СНИЛС отсутствует.'
    }),
    auth_success_snils: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Удачный вход по СНИЛС',
      description: 'СНИЛС успешно подтверждён.'
    }),
    auth_sheet_missing: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Лист не найден',
      description: 'Сервер не смог открыть лист персональных данных.'
    }),
    auth_table_empty: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Таблица пуста',
      description: 'Лист персональных данных пуст.'
    }),
    auth_structure_error: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Ошибка структуры таблицы',
      description: 'Не найден обязательный столбец.'
    }),
    auth_config_error: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Ошибка конфигурации столбцов',
      description: 'Ошибка серверной настройки столбцов авторизации.'
    }),

    warning: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Предупреждение: {text}',
      description: 'Пользователю показано предупреждение.'
    }),
    error: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Ошибка: {text}',
      description: 'Пользователю показана ошибка.'
    }),

    section_visit: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Перешёл в раздел {section}',
      description: 'Пользователь переключил верхний раздел сайта.'
    }),
    fill_form_click: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Заполнить форму',
      description: 'Нажатие кнопки подготовки формы посещаемости.'
    }),

    edit_click: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Редактировать — {field}',
      description: 'Вход в режим редактирования поля.'
    }),
    cancel_edit: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Отмена — {field}',
      description: 'Выход из редактирования без серверного сохранения.'
    }),
    save_click: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Сохранить — {field}',
      description: 'Нажатие кнопки сохранения поля.'
    }),
    draft_saved: Object.freeze({
      enabled: true,
      type: 'local',
      template: 'Локально сохранено: {field} → {value}',
      description: 'Изменение поля сохранено в локальный черновик после blur.'
    }),
    save_success: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Сохранено на сервере: {field} → {value}',
      description: 'Сервер подтвердил сохранение поля.'
    }),
    save_error: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Ошибка сохранения: {field}',
      description: 'Сервер не подтвердил сохранение поля.'
    }),

    file_open: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Открыл файл: {field}',
      description: 'Пользователь открыл прикреплённый файл.'
    }),
    file_attach_click: Object.freeze({
      enabled: true,
      type: 'action',
      template: 'Нажал: Прикрепить файл — {field}',
      description: 'Нажатие кнопки прикрепления файла.'
    }),
    file_upload_success: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Файл загружен: {field}',
      description: 'Сервер подтвердил загрузку файла.'
    }),
    file_upload_error: Object.freeze({
      enabled: true,
      type: 'result',
      template: 'Ошибка загрузки файла: {field}',
      description: 'Сервер не подтвердил загрузку файла.'
    })
  })
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
