/*************************************************
 * ДОКУМЕНТЫ
 *************************************************/
const DOCUMENTS_CONFIG = Object.freeze({
  // Ссылка на корневую папку для прикрепляемых файлов:
  attachmentsFolderUrl: 'https://drive.google.com/drive/folders/1AyjWNspWbBVswPdrSy0M-JEbvZBzsjq1',
  // Название столбца, куда записывать дату обновления информации о школе:
  schoolInfoUpdatedHeader: 'Дата обн. инф. о школе (С)',
  // Название папки пользователя:
  userFolder: Object.freeze({
    template: 'Фамилия Имя Отчество (С) - Дата рождения (С) - Год набора',
    headers: Object.freeze(['Фамилия Имя Отчество (С)', 'Дата рождения (С)', 'Год набора'])
  }),
  // Настройки отдельных файлов:
  files: Object.freeze({
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
  }),
  // Публичные настройки документов:
  client: Object.freeze({
    timings: Object.freeze({
      // Время действия временной ссылки на открываемый файл, миллисекунд:
      attachmentObjectUrlLifetimeMs: 60 * 1000
    }),
    documentSections: Object.freeze({
      parentSuffixes: Object.freeze(['П', 'М', 'Д']),
      parentFields: Object.freeze(['Фамилия Имя Отчество', 'Телефон +7', 'Электронная почта', 'Дата рождения', 'Паспорт: Серия, номер', 'Паспорт: Кем выдан', 'Паспорт: Когда выдан', 'Паспорт: Прописка', 'Паспорт: Код подразделения', 'Марка автомобиля', 'гос. номер автомобиля']),
      starts: Object.freeze({ 'Фамилия Имя Отчество (С)': 'Спортсмен', 'Фамилия Имя Отчество (П)': 'Отец', 'Фамилия Имя Отчество (М)': 'Мать', 'Фамилия Имя Отчество (Д)': 'Другой законный представитель' })
    })
  })
});

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
  // YEAR — 4 цифры в настроенном диапазоне года набора.
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
    YEAR: { title: 'Год', placeholder: '2024', regex: '^\\d{4}$', special: 'year', min: 1950, maxOffset: 1 },
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
