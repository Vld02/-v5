/*************************************************
 * ПОСЕЩАЕМОСТЬ И ТРЕНИРОВКИ
 *************************************************/
const TRAINING_CONFIG = Object.freeze({
  // Ссылка на таблицу с тренировками:
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1K1TtjIL2retzFoXBlQaePKbeKIEkMZZedZX-Ans4VjY/edit?usp=sharing',
  // Название листа с заполненными тренировками:
  responseSheetName: 'ОтветыV5',
  // Заголовки таблицы тренировок:
  headers: Object.freeze({ timestamp: 'Отметка времени', date: 'Дата тренировки', coach: 'Тренер присутствовал:', place: 'Место проведения занятия' }),
  // Первый столбец с ФИО выбранных спортсменов:
  responseFioStartColumn: 12,
  // Последний столбец с ФИО выбранных спортсменов:
  responseFioEndColumn: 32,
  // Столбец с ФИО спортсмена:
  athleteNameHeader: 'Фамилия Имя Отчество (С)',
  // Столбец с тренировочной группой:
  athleteGroupHeader: 'Тренировочная группа',
  // Надпись для спортсмена без группы:
  noGroupLabel: 'Без группы',
  // Ключ кэша списка ФИО для посещаемости. Можно указать любое уникальное значение:
  NAMES_CACHE_KEY: 'dbv5_full_names_v1',
  // Время хранения списка ФИО для посещаемости в кэше, секунд:
  NAMES_CACHE_TTL_SECONDS: 300,
  // Подбор сокращённого ФИО:
  maxNameMatchErrors: 2,
  nameMatchOptionsLimit: 3
});

// Публичные настройки посещаемости и тренировок:
const TRAINING_CLIENT_CONFIG = Object.freeze({
  storageKeys: Object.freeze({
    attendanceRows: 'attendanceRowsDraftV1', attendanceDraftDeleteAfter: 'attendanceDraftDeleteAfterAtV1',
    lastSilentSync: 'lastSilentSyncAt', trainingHistory: 'trainingHistoryItemsV1'
  }),
  timings: Object.freeze({
    attendanceDraftTtlMs: 30 * 60 * 1000, silentSyncIntervalMs: 15 * 60 * 1000,
    nameMatchDebounceMs: 400, attendanceNoticeMs: 1800, trainingSyncSuccessMs: 950
  }),
  urls: Object.freeze({
    trainingSheet: TRAINING_CONFIG.spreadsheetUrl,
    attendanceTable: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQAnNOAevcu7f79FGp8ol6XHXki2BUa_zXnujvbk-g3EzvQBXkVqFuK-SKMTfDHhMlRikuu220nf77D/pubhtml?gid=866330013&single=true&widget=true&headers=false',
    attendanceForm: 'https://docs.google.com/forms/d/e/1FAIpQLSdGwQQhPaY3wXjT90TX2daQx7U-mjnfkoL_7VZ9nJ8NUqbKtw/viewform?entry.286976530'
  }),
  ui: Object.freeze({ initialTrainingVisibleCount: 5, attendanceIframeTitle: 'Таблица посещаемости' })
});
