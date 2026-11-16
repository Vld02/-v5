/*************************************************
 * ВНУТРЕННИЕ ТЕХНИЧЕСКИЕ НАСТРОЙКИ
 *************************************************/
const TechnicalConfig = Object.freeze({
  // Ключ, по которому в ScriptCache сохраняется список ФИО спортсменов.
  // Изменение создаёт отдельную техническую запись кэша:
  NAMES_CACHE_KEY: 'dbv5_full_names_v1',

  // Внутренние имена записей, сохраняемых сайтом в памяти браузера.
  // Переименование каждого ключа создаёт новое локальное хранилище данных:
  storageKeys: Object.freeze({
    attendanceRows: 'attendanceRowsDraftV1',
    attendanceDraftDeleteAfter: 'attendanceDraftDeleteAfterAtV1',
    lastSilentSync: 'lastSilentSyncAt',
    trainingHistory: 'trainingHistoryItemsV1',
    documentEditDrafts: 'documentEditDraftsV1'
  })
});
