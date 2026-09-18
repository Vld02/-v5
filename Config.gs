/*************************************************
 * ЕДИНЫЙ КОНФИГ ПРИЛОЖЕНИЯ
 *
 * Меняйте рабочие параметры только в этом файле.
 *
 * БЫСТРАЯ ПАМЯТКА ПО ИЗМЕНЕНИЮ
 * 1. Изменяйте значение справа от двоеточия, а название слева НЕ меняйте:
 *    его использует программный код.
 * 2. Строки всегда заключайте в одинарные кавычки: 'Результат'. Числа
 *    пишите без кавычек: 300. true включает параметр, false выключает.
 * 3. После сохранения создайте новую версию и разверните web-app заново.
 * 4. Если изменили имя листа или заголовок столбца, сначала переименуйте
 *    его в Google Sheets, затем укажите точно такое же имя здесь.
 *
 * СОДЕРЖАНИЕ:
 *   WEBAPP_FAVICON_URL — иконка вкладки;
 *   CONFIG             — таблица, папка, даты, кэш, вложения;
 *   APP_CONFIG         — название приложения и подписи журнала;
 *   LOG_CONFIG         — состав и время объединения записей журнала;
 *   EDIT_CONFIG        — правила ввода и карта доступа к столбцам;
 *   CLIENT_CONFIG      — публичные настройки, которые получает браузер.
 *
 * ВНИМАНИЕ О БЕЗОПАСНОСТИ: всё из CLIENT_CONFIG посетитель может увидеть
 * в исходном коде страницы. Пароли, ключи API, токены и приватные ссылки
 * туда не добавляйте. Их нужно хранить в Script Properties.
 * Code.gs использует серверные значения, а CLIENT_CONFIG передаётся
 * в index.html при каждом открытии web-app.
 *************************************************/

/*************************************************
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 *************************************************/
// Публичный HTTPS-адрес PNG или ICO. Пример: 'https://site.ru/icon.png'.
// Пустая строка отключит отдельную иконку именно в обёртке Apps Script.
const WEBAPP_FAVICON_URL = 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.ico';

// Серверные параметры. Object.freeze защищает их от случайного изменения кодом.
const CONFIG = Object.freeze({
  // ID находится между /d/ и /edit в URL таблицы Google Sheets. Вставьте только ID.
  SPREADSHEET_ID: '1PITVXQ48g0hwtx4YSWB7OOy37zvujj9hhts-7eGR1aQ',
  // Лист с персональными данными; значение — точное имя вкладки Sheets.
  RESULT_SHEET_NAME: 'Результат',
  // Лист журнала; если его нет, приложение создаст его с этим именем.
  LOG_SHEET_NAME: 'Входы',
  // Цвет заголовка разрешённого к показу столбца. Обычно '#ffff00' (жёлтый).
  YELLOW: '#ffff00',
  // Часовой пояс Apps Script: например 'GMT+3', 'Europe/Moscow' или 'UTC'.
  TIMEZONE: 'GMT+3',
  // Формат дат Utilities.formatDate: 'dd.MM.yyyy', 'yyyy-MM-dd' и т. п.
  DATE_FORMAT: 'dd.MM.yyyy',
  // Произвольный уникальный ключ кэша; измените, чтобы принудительно сбросить кэш ФИО.
  NAMES_CACHE_KEY: 'dbv5_full_names_v1',
  // Время жизни кэша в секундах: 60 = 1 минута, 300 = 5 минут, 3600 = 1 час.
  NAMES_CACHE_TTL_SECONDS: 300,
  // Границы допустимого года набора: текущий год + смещение.
  ENROLLMENT_YEAR_MIN: 1950,
  ENROLLMENT_YEAR_OFFSET: 1,
  // Настройка входа. Значения — точные заголовки столбцов листа RESULT_SHEET_NAME.
  // loginHeader — ФИО для входа, passwordHeader — дата рождения, snilsHeader — СНИЛС.
  AUTH: Object.freeze({
    loginHeader: 'Фамилия Имя Отчество (С)',
    passwordHeader: 'Дата рождения (С)',
    snilsHeader: 'Снилс: номер'
  }),
  // Источник истории тренировок. responseFioStartColumn/endColumn — индексы с нуля:
  // 12 = столбец M, 32 = AG. Из этого диапазона читаются выбранные ФИО формы.
  TRAINING: Object.freeze({
    spreadsheetId: '1K1TtjIL2retzFoXBlQaePKbeKIEkMZZedZX-Ans4VjY',
    responseSheetName: 'ОтветыV5',
    headers: Object.freeze({ timestamp: 'Отметка времени', date: 'Дата тренировки', coach: 'Тренер присутствовал:', place: 'Место проведения занятия' }),
    responseFioStartColumn: 12,
    responseFioEndColumn: 32,
    athleteNameHeader: 'Фамилия Имя Отчество (С)',
    athleteGroupHeader: 'Тренировочная группа',
    noGroupLabel: 'Без группы'
  }),
  // Небольшие серверные интервалы/лимиты. lockWaitMs — ожидание записи в журнал;
  // maxNameMatchErrors — допустимые опечатки в сокращённом ФИО; nameMatchOptionsLimit — число вариантов.
  OPERATIONS: Object.freeze({ lockWaitMs: 1000, maxNameMatchErrors: 2, nameMatchOptionsLimit: 3 }),
  // Служебные заголовки, используемые автоматическими действиями после сохранения.
  FIELDS: Object.freeze({ schoolInfoUpdatedHeader: 'Дата обн. инф. о школе (С)' }),
  // Корневая папка «Пользователи». Внутри неё создаётся папка для каждой строки.
  // ID корневой папки Drive из URL .../folders/ID. Не URL и не название папки.
  ATTACHMENTS_FOLDER_ID: '1AyjWNspWbBVswPdrSy0M-JEbvZBzsjq1',
  // Шаблоны используют значения столбцов листа «Результат».
  // template — строка имени. В ней указывайте заголовки точно как в headers,
  // разделяя их текстом (например, 'ФИО - Дата'). headers — массив используемых
  // заголовков. Добавьте/удалите оба одновременно. Ключ FILES — это точный
  // заголовок FILE-столбца из EDIT_CONFIG.fields.
  ATTACHMENT_NAMING: Object.freeze({
    USER_FOLDER: Object.freeze({
      template: 'Фамилия Имя Отчество (С) - Дата рождения (С) - Год набора',
      headers: Object.freeze(['Фамилия Имя Отчество (С)', 'Дата рождения (С)', 'Год набора'])
    }),
    FILES: Object.freeze({
      'Свидетельство: скан (С)': Object.freeze({ template: 'Свидетельство: скан (С) - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Паспорт: скан (С)': Object.freeze({ template: 'Паспорт: скан (С) - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Снилс: Скан': Object.freeze({ template: 'Снилс: Скан - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Полис: Скан': Object.freeze({ template: 'Полис: Скан - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Страховка: Скан': Object.freeze({ template: 'Страховка: Скан - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Мед допуск: Скан': Object.freeze({ template: 'Мед допуск: Скан - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Русада: Скан': Object.freeze({ template: 'Русада: Скан - Фамилия Имя Отчество (С)', headers: Object.freeze(['Фамилия Имя Отчество (С)']) }),
      'Паспорт: Скан (П)': Object.freeze({ template: 'Паспорт: Скан (П) - Фамилия Имя Отчество (П)', headers: Object.freeze(['Фамилия Имя Отчество (П)']) }),
      'Паспорт: Скан (М)': Object.freeze({ template: 'Паспорт: Скан (М) - Фамилия Имя Отчество (М)', headers: Object.freeze(['Фамилия Имя Отчество (М)']) }),
      'Паспорт: Скан (Д)': Object.freeze({ template: 'Паспорт: Скан (Д) - Фамилия Имя Отчество (Д)', headers: Object.freeze(['Фамилия Имя Отчество (Д)']) })
    })
  })
});


/**
 * Отображаемое название приложения и названия разделов в журнале.
 * APP_TITLE используется в заголовке вкладки и web-app. SECTION_NAMES
 * сопоставляет технический ключ вкладки с понятной записью в журнале.
 */
const APP_CONFIG = Object.freeze({
  // Текст на вкладке браузера и в заголовке web-app. Любая короткая строка.
  APP_TITLE: 'ДБВv5',
  // Цвет браузерной темы в #RRGGBB; используйте любой корректный CSS-цвет.
  THEME_COLOR: '#179bcf',
  // Публичные изображения для вкладок и ярлыка на устройстве. Полные HTTPS-URL.
  ICON_URLS: Object.freeze({
    icon32: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/32.png',
    icon72: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/72.png',
    icon192: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/192.png',
    icon512: 'https://raw.githubusercontent.com/Vld02/-v5/refs/heads/main/512.png'
  }),
  SECTION_NAMES: Object.freeze({
    // Ключи docs / attendance / gear не менять: это ключи вкладок в коде.
    docs: 'ДБВv5 Документы',
    attendance: 'ДБВv5 Посещаемость',
    gear: 'ДБВv5 Снаряжение'
  })
});

/**
 * Параметры журнала действий. LOG_COLUMNS определяет и порядок, и
 * названия колонок: при изменении существующий заголовок будет обновлён.
 * MAX_AGE_MINUTES — окно, в котором открытие страницы и вход объединяются
 * в одну многострочную запись.
 */
const LOG_CONFIG = Object.freeze({
  // Порядок обязателен: значения лога записываются в той же последовательности.
  COLUMNS: Object.freeze([
    'Дата/время входа',
    'Логин',
    'Пароль',
    'СНИЛС',
    'IP',
    'Устройство',
    'Браузер',
    'Статус входа'
  ]),
  // Число минут: 0 — не объединять записи, 30 — стандартное окно, 60 — час.
  MAX_AGE_MINUTES: 30
});
const LOG_COLUMNS = LOG_CONFIG.COLUMNS;
const LOG_MAX_AGE_MINUTES = LOG_CONFIG.MAX_AGE_MINUTES;

/* ============================================================
   НАБОР ПРАВИЛ РЕДАКТИРОВАНИЯ — СЕРВЕРНЫЙ ИСТОЧНИК ПРОВЕРКИ

   Серверная копия той же архитектуры, которую использует браузер
   для удобства ввода. В Apps Script HTML и Code.gs выполняются в
   разных контекстах, поэтому сервер обязан иметь собственный доступ
   к конфигурации и никогда не доверять объектам, изменённым в браузере.

   ВАЖНО: именно этот блок защищает запись в Google Таблицу. Перед
   setValue() updateResultCell() проверяет существование поля,
   editable, required и regex по EDIT_CONFIG.fields + EDIT_CONFIG.rules.
   ============================================================ */
const EDIT_CONFIG = Object.freeze({
  // КАТАЛОГ ВАРИАНТОВ rule:
  // TEXT — любой текст; SUGGEST_TEXT — любой текст с подсказками из suggestions.
  // FULL_NAME_RU — три русских слова: «Иванов Иван Иванович».
  // YEAR — 4 цифры в диапазоне CONFIG.ENROLLMENT_YEAR_MIN..текущий год + OFFSET.
  // CLASS_COURSE — 0–11 либо римские I–VI; RU_UPPER_LETTER — одна А–Я/Ё.
  // PHONE_RU — строго «+7 999 123-45-67»; EMAIL — адрес с @ и доменом.
  // CERTIFICATE_RU — «IV-АБ № 123456»; DATE_RU — «ДД.ММ.ГГГГ» с реальной датой.
  // SNILS — 11 цифр, дефисы/пробелы допускаются; PASSPORT_RU — «12 34 567890».
  // PASSPORT_DIVISION_CODE — «123-456»; MED_POLICY_NUMBER — 16 цифр группами 4.
  // MGFSO_ID — ровно 7 цифр; FILE — загрузка файла, текст вручную не принимается.
  rules: Object.freeze({
    // TEXT: Любая последовательность символов, включая цифры и пробелы.
    TEXT: { title: 'Текст', placeholder: '', regex: '^.*$', special: '' },
    // SUGGEST_TEXT: Текст можно ввести вручную или выбрать подсказку; обязательно задайте suggestions у поля.
    SUGGEST_TEXT: { title: 'Текст из подсказок', placeholder: '', regex: '^.*$', special: 'suggest' },
    // FULL_NAME_RU: Ровно три слова с русской заглавной буквы: фамилия, имя и отчество.
    FULL_NAME_RU: { title: 'ФИО', placeholder: 'Иванов Иван Иванович', regex: '^\\s*[А-ЯЁ][а-яё]+\\s+[А-ЯЁ][а-яё]+\\s+[А-ЯЁ][а-яё]+\\s*$', special: 'fullname' },
    // YEAR: Ровно четыре цифры в настроенном диапазоне года набора.
    YEAR: { title: 'Год', placeholder: '2024', regex: '^\\d{4}$', special: 'year' },
    // CLASS_COURSE: Арабское число 0–11 или римское обозначение I, II, III, IV, V, VI.
    CLASS_COURSE: { title: 'Класс / курс', placeholder: '7', regex: '^(?:[0-9]|1[01]|I|II|III|IV|V|VI)$', special: 'classCourse' },
    // RU_UPPER_LETTER: Одна заглавная русская буква, включая Ё.
    RU_UPPER_LETTER: { title: 'Русская заглавная буква', placeholder: 'А', regex: '^[А-ЯЁ]$', special: 'singleRuUpper' },
    // PHONE_RU: Только формат +7 999 123-45-67; пробелы и дефисы обязательны.
    PHONE_RU: { title: 'Номер телефона', placeholder: '+7 999 123-45-67', regex: '^\\+7\\s\\d{3}\\s\\d{3}-\\d{2}-\\d{2}$', special: 'phoneRu' },
    // EMAIL: Обычный e-mail с @ и доменной частью после точки.
    EMAIL: { title: 'Электронная почта', placeholder: 'name@example.ru', regex: '^[A-Za-z0-9.!#$%&\'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\\.[A-Za-z0-9-]+)+$', special: 'email' },
    // CERTIFICATE_RU: Серия римскими цифрами, русские буквы, знак № и шесть цифр.
    CERTIFICATE_RU: { title: 'Свидетельство о рождении', placeholder: 'IV-АБ № 123456', regex: '^([VIX]{1,4}-[А-ЯЁ]{1,3}\\s*№\\s*\\d{6})$', special: 'certificate' },
    // DATE_RU: Формат ДД.ММ.ГГГГ; дополнительно проверяется существование даты.
    DATE_RU: { title: 'Дата', placeholder: 'ДД.ММ.ГГГГ', regex: '^\\d{2}\\.\\d{2}\\.\\d{4}$', special: 'date' },
    // SNILS: 11 цифр; допускаются дефисы или пробелы между группами.
    SNILS: { title: 'СНИЛС', placeholder: '000-000-000-00', regex: '^\\d{3}[-\\s]?\\d{3}[-\\s]?\\d{3}[-\\s]?\\d{2}$', special: 'snils' },
    // PASSPORT_RU: Две цифры, пробел, две цифры, пробел, шесть цифр.
    PASSPORT_RU: { title: 'Паспорт РФ', placeholder: '12 34 567890', regex: '^\\d{2}\\s\\d{2}\\s\\d{6}$', special: 'passportRu' },
    // PASSPORT_DIVISION_CODE: Три цифры, дефис, три цифры.
    PASSPORT_DIVISION_CODE: { title: 'Код подразделения', placeholder: '123-456', regex: '^\\d{3}-\\d{3}$', special: 'passportDivisionCode' },
    // MED_POLICY_NUMBER: 16 цифр, разделённые пробелами на четыре группы по четыре.
    MED_POLICY_NUMBER: { title: 'Номер медполиса', placeholder: '1234 5678 9012 3456', regex: '^\\d{4}\\s\\d{4}\\s\\d{4}\\s\\d{4}$', special: 'medPolicyNumber' },
    // MGFSO_ID: Только семь цифр.
    MGFSO_ID: { title: 'ID МГФСО', placeholder: '1234567', regex: '^\\d{7}$', special: 'mgfsoId' },
    // FILE: Кнопка прикрепления файла; ручной текст для ячейки запрещён.
    FILE: { title: 'Файл', placeholder: '', regex: '^$', special: 'file' }
  }),

  /* ============================================================
   СООТВЕТСТВИЕ ПОЛЕЙ И ПРАВИЛ — СЕРВЕРНАЯ КАРТА ДОСТУПА

   КАЖДАЯ строка ниже — разрешение на редактирование одного столбца.
   Ключ слева — ТОЧНЫЙ заголовок в первой строке листа «Результат».
   Если его нет в карте, запись в него запрещена — это безопасное значение
   по умолчанию. Чтобы добавить столбец, скопируйте любую строку, замените
   ключ и подберите rule из полного списка выше.

   ПАРАМЕТРЫ КАЖДОЙ СТРОКИ:
   - editable: true — пользователь может сохранить значение; false — поле
     показывается, но сервер запрещает его изменение.
   - rule: один из TEXT, SUGGEST_TEXT, FULL_NAME_RU, YEAR, CLASS_COURSE,
     RU_UPPER_LETTER, PHONE_RU, EMAIL, CERTIFICATE_RU, DATE_RU, SNILS,
     PASSPORT_RU, PASSPORT_DIVISION_CODE, MED_POLICY_NUMBER, MGFSO_ID, FILE.
     Полное объяснение каждого варианта находится непосредственно над картой.
   - required: true — пустое значение нельзя сохранить; false — очистка поля
     разрешена. required не отменяет проверку rule.
   - isIdentityField: true — изменение обновляет текущие данные входа
     пользователя; используйте только для ФИО спортсмена, даты рождения и СНИЛС.
   - description — подсказка пользователю; example — пример корректного ввода.
   - suggestions — только для SUGGEST_TEXT: sourceSheet — лист-источник,
     sourceHeader — его точный заголовок, startRow — первая строка данных
     (обычно 2, если первая строка содержит заголовки).
   - onSave: 'UPDATE_SCHOOL_DATE' — после сохранения школы выполняет
     встроенное действие обновления даты школы. Не указывайте другое значение:
     других действий в коде нет.

   Существующие строки намеренно содержат только допустимые варианты. Для
   запрета уже существующего поля смените editable: true на editable: false;
   для полного запрета удалите строку из карты.
   ============================================================ */
  fields: Object.freeze({
    'Фамилия Имя Отчество (С)': { editable: true, rule: 'FULL_NAME_RU', required: true, isIdentityField: true, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Дата рождения (С)': { editable: true, rule: 'DATE_RU', required: true, isIdentityField: true, description: 'Дата рождения в формате ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Месяц рождения (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Год набора': { editable: true, rule: 'YEAR', required: false, isIdentityField: false, description: 'Год в динамическом диапазоне 1950 — текущий год + 1.', example: '2024' },
    'Пол (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Школа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, onSave: 'UPDATE_SCHOOL_DATE', description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Школы', startRow: 2 } },
    'Класс / курс': { editable: true, rule: 'CLASS_COURSE', required: false, isIdentityField: false, description: '0-11 или I-VI.', example: '7' },
    'Литера класса (буква)': { editable: true, rule: 'RU_UPPER_LETTER', required: false, isIdentityField: false, description: 'Одна заглавная русская буква.', example: 'А' },
    'Директор школы: Фамилия Имя Отчество': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Адрес регистрации / Прописка (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Телефон +7 (С)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (С)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Фамилия Имя Отчество (П)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (П)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (П)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (П)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (П)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (С)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (С)', startRow: 2 } },
    'Паспорт: Кем выдан (П)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (П)', startRow: 2 } },
    'Паспорт: Когда выдан (П)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (П)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (П)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Фамилия Имя Отчество (М)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (М)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (М)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (М)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (М)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (М)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (М)', startRow: 2 } },
    'Паспорт: Когда выдан (М)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (М)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (М)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Фамилия Имя Отчество (Д)': { editable: true, rule: 'FULL_NAME_RU', required: false, isIdentityField: false, description: 'Фамилия, имя и отчество.', example: 'Иванов Иван Иванович' },
    'Телефон +7 (Д)': { editable: true, rule: 'PHONE_RU', required: false, isIdentityField: false, description: 'Российский номер +7 999 123-45-67.', example: '+7 999 123-45-67' },
    'Электронная почта (Д)': { editable: true, rule: 'EMAIL', required: false, isIdentityField: false, description: 'Адрес электронной почты.', example: 'name@example.ru' },
    'Дата рождения (Д)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Серия, номер (Д)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Паспорт: Кем выдан (Д)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан (Д)', startRow: 2 } },
    'Паспорт: Когда выдан (Д)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Прописка (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Паспорт: Код подразделения (Д)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Марка автомобиля (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'гос. номер автомобиля (Д)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Свидетельство: Серия, номер (С)': { editable: true, rule: 'CERTIFICATE_RU', required: false, isIdentityField: false, description: 'Серия и номер свидетельства.', example: 'IV-АБ № 123456' },
    'Паспорт: Серия, номер (С)': { editable: true, rule: 'PASSPORT_RU', required: false, isIdentityField: false, description: 'Серия и номер паспорта в формате 12 34 567890.', example: '12 34 567890' },
    'Свидетельство: скан (С)': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан свидетельства о рождении.' },
    'Паспорт: скан (С)': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан паспорта спортсмена.' },
    'Снилс: Скан': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан СНИЛС.' },
    'Полис: Скан': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан медицинского полиса.' },
    'Страховка: Скан': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан страховки.' },
    'Мед допуск: Скан': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан медицинского допуска.' },
    'Русада: Скан': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан документа РУСАДА.' },
    'Паспорт: Скан (П)': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан паспорта отца.' },
    'Паспорт: Скан (М)': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан паспорта матери.' },
    'Паспорт: Скан (Д)': { editable: true, rule: 'FILE', required: false, isIdentityField: false, description: 'Скан паспорта другого законного представителя.' },
    'Свидетельство: Кем выдан (С)': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Свидетельство: Кем выдан (С)', startRow: 2 } },
    'Паспорт: Кем выдан': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Паспорт: Кем выдан', startRow: 2 } },
    'Паспорт или Свидетельство: Кем выдан (С)': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'Свидетельство: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт или Свидетельство: Когда выдан (С)': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Паспорт: Код подразделения (С)': { editable: true, rule: 'PASSPORT_DIVISION_CODE', required: false, isIdentityField: false, description: 'Код подразделения в формате 123-456.', example: '123-456' },
    'Снилс: номер': { editable: true, rule: 'SNILS', required: false, isIdentityField: true, description: 'СНИЛС из 11 цифр.', example: '000-000-000-00' },
    'Полис: Страховая компания': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Результат', sourceHeader: 'Полис: Страховая компания', startRow: 2 } },
    'Тренировочная группа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Группы тренировки', startRow: 5 } },
    'МГФСО группа': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Группы МГФСО', startRow: 2 } },
    'Тренер МГФСО': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Тренер МГФСО', startRow: 2 } },
    'Разряд': { editable: true, rule: 'SUGGEST_TEXT', required: false, isIdentityField: false, description: 'Произвольный текст с подсказками.', example: 'Значение из списка', suggestions: { sourceSheet: 'Списки данных', sourceHeader: 'Список разрядов', startRow: 2 } },
    'Полис: Номер': { editable: true, rule: 'MED_POLICY_NUMBER', required: false, isIdentityField: false, description: 'Номер медицинского полиса в формате 1234 5678 9012 3456.', example: '1234 5678 9012 3456' },
    'Дата зачисления в МГФСО': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Укажите дату зачисления в МГФСО.', example: '15.08.2024' },
    'Дата получения разряда': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Страховка: Действительна до': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'Страховка: Номер и компаия': { editable: true, rule: 'TEXT', required: false, isIdentityField: false, description: 'Произвольное текстовое значение.', example: 'Текст' },
    'ID номер МГФСО': { editable: true, rule: 'MGFSO_ID', required: false, isIdentityField: false, description: 'ID МГФСО из 7 цифр.', example: '1234567' },
    'Мед допуск до': { editable: true, rule: 'DATE_RU', required: false, isIdentityField: false, description: 'Дата ДД.ММ.ГГГГ.', example: '01.09.2010' },
    'РУСАДА': { editable: true, rule: 'YEAR', required: false, isIdentityField: false, description: 'Год в динамическом диапазоне 1950 — текущий год + 1.', example: '2024' }
  })
});


/**
 * Параметры браузера, которые разрешено менять без поиска по HTML.
 * Значения передаются посетителю, поэтому не добавляйте сюда секреты.
 *
 * storageKeys: ключи localStorage; изменяйте их только если намеренно
 * хотите начать новую локальную историю/черновик у пользователей.
 * attendanceFormUrl: URL Google Forms БЕЗ параметра entry — он добавляется
 * автоматически вместе со списком выбранных ФИО.
 */
const CLIENT_CONFIG = Object.freeze({
  // Ключи браузерного localStorage. Оставьте как есть, чтобы не потерять
  // сохранённые у пользователей черновики. Новое уникальное имя начнёт чистое хранилище.
  storageKeys: Object.freeze({
    // Сохраняемые данные входа. Менять только для принудительного сброса устройств.
    savedLogin: 'savedLogin',
    savedDate: 'savedDate',
    savedSnils: 'savedSnils',
    isLoggedIn: 'isLoggedIn',
    activeTab: 'activeTab',
    attendanceRows: 'attendanceRowsDraftV1',
    attendanceDraftDeleteAfter: 'attendanceDraftDeleteAfterAtV1',
    lastSilentSync: 'lastSilentSyncAt',
    trainingHistory: 'trainingHistoryItemsV1'
  }),
  // Все значения в миллисекундах: 1000 = 1 секунда, 60000 = 1 минута.
  timings: Object.freeze({
    // Сколько хранить черновик посещаемости после открытия формы.
    attendanceDraftTtlMs: 30 * 60 * 1000,
    // Минимальный интервал фоновой синхронизации авторизованного пользователя.
    silentSyncIntervalMs: 15 * 60 * 1000,
    // Через сколько скрывать обычное всплывающее сообщение; 0 не используйте.
    toastAutoHideMs: 3500,
    // Скорость анимации точек загрузки, задержка скрытия подсказок и уведомлений.
    authLoadingAnimationMs: 450,
    suggestionBlurDelayMs: 120,
    nameMatchDebounceMs: 400,
    attendanceNoticeMs: 1800,
    trainingSyncSuccessMs: 950,
    // Как долго открытая ссылка Blob на прикреплённый файл остаётся действительной.
    attachmentObjectUrlLifetimeMs: 60 * 1000
  }),
  // Только публичные HTTPS-ссылки. Они видны каждому посетителю страницы.
  urls: Object.freeze({
    // Сервис, возвращающий JSON вида { ip: '...' }; можно заменить совместимым API.
    ipLookup: 'https://api.ipify.org?format=json',
    // Ссылка «открыть таблицу истории тренировок»; вставьте полный URL таблицы.
    trainingSheet: 'https://docs.google.com/spreadsheets/d/1K1TtjIL2retzFoXBlQaePKbeKIEkMZZedZX-Ans4VjY/edit?usp=sharing',
    // URL формы ДО значения: оставьте параметр entry.<ID> без знака '=' в конце.
    // Например: .../viewform?entry.123456. Приложение добавит '=ФИО%0AФИО'.
    // Опубликованная таблица для встроенного режима посещаемости.
    attendanceTable: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQAnNOAevcu7f79FGp8ol6XHXki2BUa_zXnujvbk-g3EzvQBXkVqFuK-SKMTfDHhMlRikuu220nf77D/pubhtml?gid=866330013&single=true&widget=true&headers=false',
    attendanceForm: 'https://docs.google.com/forms/d/e/1FAIpQLSdGwQQhPaY3wXjT90TX2daQx7U-mjnfkoL_7VZ9nJ8NUqbKtw/viewform?entry.286976530'
  }),
  // Эти значения дублируют серверную проверку только для удобства ввода.
  // Реальная защита остаётся на сервере в CONFIG.
  // Заголовки полей авторизации, доступные браузеру для обновления формы после сохранения.
  // Не меняйте отдельно от CONFIG.AUTH: значения должны совпадать.
  authFields: Object.freeze({ login: CONFIG.AUTH.loginHeader, password: CONFIG.AUTH.passwordHeader, snils: CONFIG.AUTH.snilsHeader }),
  ui: Object.freeze({
    // Сколько тренировок показывать первоначально; целое число не меньше 1.
    initialTrainingVisibleCount: 5,
    attendanceIframeTitle: 'Таблица посещаемости'
  }),
  validation: Object.freeze({
    enrollmentYearMin: CONFIG.ENROLLMENT_YEAR_MIN,
    enrollmentYearOffset: CONFIG.ENROLLMENT_YEAR_OFFSET
  }),
  // Структура карточки «Документы». Изменяйте только если одновременно
  // меняете соответствующие заголовки в таблице.
  documentSections: Object.freeze({
    // Суффиксы законных представителей: П — отец, М — мать, Д — другой представитель.
    parentSuffixes: Object.freeze(['П', 'М', 'Д']),
    // Общие части заголовков, по которым определяется заполненность блока представителя.
    parentFields: Object.freeze([
      'Фамилия Имя Отчество', 'Телефон +7', 'Электронная почта', 'Дата рождения',
      'Паспорт: Серия, номер', 'Паспорт: Кем выдан', 'Паспорт: Когда выдан',
      'Паспорт: Прописка', 'Паспорт: Код подразделения', 'Марка автомобиля',
      'гос. номер автомобиля'
    ]),
    // Заголовок первого поля раздела => видимое название этого раздела на карточке.
    starts: Object.freeze({
      'Фамилия Имя Отчество (С)': 'Спортсмен',
      'Фамилия Имя Отчество (П)': 'Отец',
      'Фамилия Имя Отчество (М)': 'Мать',
      'Фамилия Имя Отчество (Д)': 'Другой законный представитель'
    })
  })
});

/**
 * Возвращает JSON-конфиг как безопасный JavaScript для шаблона index.html.
 * Экранирование < предотвращает закрытие тега script значением настройки.
 * @returns {string}
 */
function getClientConfigScript() {
  return `window.CLIENT_CONFIG = ${JSON.stringify(CLIENT_CONFIG).replace(/</g, '\\u003c')};`;
}

