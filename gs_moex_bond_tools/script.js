/**
 * Константа: задержка между запросами при массовом обновлении (в миллисекундах)
 */
const DELAY_MS = 400; // 0.4 секунды

/**
 * Список всех кастомных функций для обновления
 * При добавлении новой функции просто добавьте её название в этот массив
 */
const CUSTOM_FUNCTIONS = [
  'GET_MOEX_PRICE',
  'GET_NEXT_COUPON',
  'GET_MOEX_NAME',
  'GET_COUPON_VALUE',
  'GET_MATURITY_DATE',
  'GET_NEAREST_OPTION_DATE',
];

/**
 * При открытии документа создает в меню пункт "MOEX".
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MOEX')
    .addItem('Обновить все данные (с задержкой)', 'forceRecalculatePrices')
    .addToUi();
}

/**
 * Кастомная функция для ячейки. Возвращает цену облигации по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {number | string} Последняя цена сделки или текстовое описание ошибки.
 * @customfunction
 */
function GET_MOEX_PRICE(ticker) {
  if (!ticker || ticker.trim() === '') {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cached = cache.get(ticker);
  if (cached !== null) {
    return JSON.parse(cached);
  }

  const result = fetchSinglePriceInternal(ticker);

  cache.put(ticker, JSON.stringify(result), 300); // Кэшируем результат на 5 минут

  return result;
}

/**
 * Универсальный обработчик: находит все ячейки со всеми кастомными функциями и обновляет их по очереди.
 */
function forceRecalculatePrices() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const dataRange = sheet.getDataRange();
  const allFormulas = dataRange.getFormulas();
  const targetCells = [];

  // Ищем все ячейки, содержащие любую из кастомных функций
  for (let i = 0; i < allFormulas.length; i++) {
    for (let j = 0; j < allFormulas[i].length; j++) {
      if (allFormulas[i][j]) {
        const formulaUpper = allFormulas[i][j].toUpperCase();
        // Проверяем, содержится ли хотя бы одна из функций
        const hasCustomFunction = CUSTOM_FUNCTIONS.some((fn) => formulaUpper.includes(fn));
        if (hasCustomFunction) {
          targetCells.push(sheet.getRange(i + 1, j + 1));
        }
      }
    }
  }

  if (targetCells.length === 0) {
    ui.alert(`На листе не найдено ячеек с функциями: ${CUSTOM_FUNCTIONS.join(', ')}`);
    return;
  }

  ui.alert(`Найдено ${targetCells.length} ячеек. Начинаю обновление...`, ui.ButtonSet.OK);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Обновление ${targetCells.length} ячеек...`);

  targetCells.forEach((cell) => {
    const originalFormula = cell.getFormula();
    cell.clearContent();
    SpreadsheetApp.flush();
    cell.setFormula(originalFormula);
    Utilities.sleep(DELAY_MS);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Обновление данных завершено!', 'Готово', 5);
}

/**
 * Внутренняя функция для получения данных. Возвращает цену или текст ошибки.
 * Версия 9.0: анализирует ОБА блока данных (marketdata и securities) для максимальной надежности.
 * @param {string} ticker - Торговый код бумаги.
 * @return {number | string} - Цена или текстовая ошибка.
 */
function fetchSinglePriceInternal(ticker) {
  const url = `https://iss.moex.com/iss/engines/stock/markets/bonds/securities/${encodeURIComponent(
    ticker
  )}.json?iss.meta=off`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    // Проверяем наличие ключевых блоков данных
    if (
      !data.marketdata ||
      !data.securities ||
      !data.marketdata.data ||
      !data.securities.data ||
      data.marketdata.data.length === 0 ||
      data.securities.data.length === 0
    ) {
      return `Тикер не найден или API вернул неполные данные`;
    }

    // Извлекаем данные из ОБЕИХ секций
    const marketdataColumns = data.marketdata.columns;
    const marketdataRow = data.marketdata.data[0];
    const securitiesColumns = data.securities.columns;
    const securitiesRow = data.securities.data[0];

    // Ищем цену в порядке приоритета по обоим блокам
    const price =
      marketdataRow[marketdataColumns.indexOf('LAST')] ??
      marketdataRow[marketdataColumns.indexOf('CLOSEPRICE')] ??
      securitiesRow[securitiesColumns.indexOf('PREVLEGALCLOSEPRICE')] ??
      securitiesRow[securitiesColumns.indexOf('PREVPRICE')];

    if (price === null || typeof price === 'undefined') {
      return 'Цена не найдена'; // Если ни одного значения не нашлось
    }

    return parseFloat(price);
  } catch (e) {
    return 'Ошибка скрипта';
  }
}

/**
 * Кастомная функция для ячейки. Возвращает ДАТУ СЛЕДУЮЩЕГО КУПОНА по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {Date | string} Дата следующего купона или текстовое описание ошибки.
 * @customfunction
 */
function GET_NEXT_COUPON(ticker) {
  if (!ticker || ticker.trim() === '') {
    return null;
  }

  // Используем кэш, чтобы не запрашивать одни и те же данные слишком часто
  const cache = CacheService.getScriptCache();
  const cacheKey = ticker + '_coupon'; // Уникальный ключ для кэша купонов
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    // Если дата в кэше, преобразуем ее обратно в объект Date
    const cachedValue = JSON.parse(cached);
    return cachedValue ? new Date(cachedValue) : 'Нет данных';
  }

  const result = fetchNextCouponInternal(ticker);

  // Кэшируем результат на 6 часов (21600 секунд), т.к. дата купона меняется редко
  cache.put(cacheKey, JSON.stringify(result), 21600);

  // Если результат - дата, возвращаем ее как объект Date
  if (result instanceof Date) {
    return result;
  }
  // Если результат - текст ошибки, возвращаем его
  return result;
}

/**
 * Внутренняя функция для получения данных о следующем купоне.
 * @param {string} ticker - Торговый код бумаги.
 * @return {Date | string} - Объект Date или текстовая ошибка.
 */
function fetchNextCouponInternal(ticker) {
  const url = `https://iss.moex.com/iss/engines/stock/markets/bonds/securities/${encodeURIComponent(
    ticker
  )}.json?iss.meta=off`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    // Данные о купоне лежат в блоке 'securities'
    if (
      !data.securities ||
      !data.securities.columns ||
      !data.securities.data ||
      data.securities.data.length === 0
    ) {
      return `Тикер не найден`;
    }

    const columns = data.securities.columns;
    const row = data.securities.data[0];

    const nextCouponIndex = columns.indexOf('NEXTCOUPON');

    if (nextCouponIndex === -1) {
      return 'Поле NEXTCOUPON отсутствует';
    }

    const couponDateStr = row[nextCouponIndex];

    // Проверяем, есть ли дата купона (может не быть у бумаг в обращении или погашенных)
    if (!couponDateStr || couponDateStr === '0000-00-00') {
      return 'Нет предстоящих купонов';
    }

    // Возвращаем как объект Date, чтобы Google Sheets правильно понял формат
    return new Date(couponDateStr);
  } catch (e) {
    return 'Ошибка скрипта';
  }
}

/**
 * Кастомная функция для ячейки. Возвращает НАИМЕНОВАНИЕ ОБЛИГАЦИИ по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {string} Наименование облигации или текст ошибки.
 * @customfunction
 */
function GET_MOEX_NAME(ticker) {
  if (!ticker || ticker.trim() === '') {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = ticker + '_name';
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const result = fetchBondNameInternal(ticker);

  // Кэшируем результат на 24 часа (86400 секунд), т.к. название не меняется
  cache.put(cacheKey, result, 86400);

  return result;
}

/**
 * Внутренняя функция для получения наименования облигации.
 * @param {string} ticker - Торговый код бумаги.
 * @return {string} - Наименование или текстовая ошибка.
 */
function fetchBondNameInternal(ticker) {
  const url = `https://iss.moex.com/iss/engines/stock/markets/bonds/securities/${encodeURIComponent(
    ticker
  )}.json?iss.meta=off`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    if (
      !data.securities ||
      !data.securities.columns ||
      !data.securities.data ||
      data.securities.data.length === 0
    ) {
      return `Тикер не найден`;
    }

    const columns = data.securities.columns;
    const row = data.securities.data[0];

    // Пробуем найти SECNAME (полное наименование) или SHORTNAME (краткое)
    const secNameIndex = columns.indexOf('SECNAME');
    const shortNameIndex = columns.indexOf('SHORTNAME');

    let name = null;
    if (secNameIndex !== -1) {
      name = row[secNameIndex];
    }
    if (!name && shortNameIndex !== -1) {
      name = row[shortNameIndex];
    }

    if (!name) {
      return 'Наименование не найдено';
    }

    return name;
  } catch (e) {
    return 'Ошибка скрипта';
  }
}

/**
 * Кастомная функция для ячейки. Возвращает РАЗМЕР СЛЕДУЮЩЕГО КУПОНА по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {number | string} Размер следующего купона в рублях или текст ошибки.
 * @customfunction
 */
function GET_COUPON_VALUE(ticker) {
  if (!ticker || ticker.trim() === '') {
    return null;
  }

  const cache = CacheService.getScriptCache();
  // Changed cache key to force refresh after logic update (v3)
  const cacheKey = ticker + '_coupon_value_v3';
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    return JSON.parse(cached);
  }

  const result = fetchCouponValueInternal(ticker);

  // Кэшируем результат на 6 часов (21600 секунд)
  cache.put(cacheKey, JSON.stringify(result), 21600);

  return result;
}

/**
 * Внутренняя функция для получения размера следующего купона.
 * @param {string} ticker - Торговый код бумаги.
 * @return {number | string} - Размер купона или текстовая ошибка.
 */
function fetchCouponValueInternal(ticker) {
  // 1. Пытаемся получить данные из основного источника (securities)
  const url = `https://iss.moex.com/iss/engines/stock/markets/bonds/securities/${encodeURIComponent(
    ticker
  )}.json?iss.meta=off`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    if (
      !data.securities ||
      !data.securities.columns ||
      !data.securities.data ||
      data.securities.data.length === 0
    ) {
      return `Тикер не найден`;
    }

    const columns = data.securities.columns;
    const row = data.securities.data[0];
    const couponValueIndex = columns.indexOf('COUPONVALUE');

    if (couponValueIndex !== -1) {
      const couponValue = row[couponValueIndex];
      // Если значение есть и оно валидное, возвращаем его
      // ВАЖНО: Если значение 0, считаем его отсутствующим (для флоатеров) и идем в fallback
      if (couponValue !== null && typeof couponValue !== 'undefined') {
        const val = parseFloat(couponValue);
        if (val !== 0) {
          return val;
        }
      }
    }

    // 2. Если значение купона не найдено или равно 0, пробуем альтернативный источник (bondization)
    return fetchCouponFromBondization(ticker);
  } catch (e) {
    return 'Ошибка скрипта: ' + e.message;
  }
}

/**
 * Кастомная функция для ячейки. Возвращает ДАТУ ПОГАШЕНИЯ облигации по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {Date | string} Дата погашения или текстовое описание ошибки.
 * @customfunction
 */
function GET_MATURITY_DATE(ticker) {
  if (!ticker || ticker.trim() === '') {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = ticker + '_maturity';
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    const cachedValue = JSON.parse(cached);
    return cachedValue ? new Date(cachedValue) : 'Нет данных';
  }

  const result = fetchMaturityDateInternal(ticker);

  // Кэшируем результат на 24 часа (86400 секунд), т.к. дата погашения не меняется
  cache.put(cacheKey, JSON.stringify(result), 86400);

  if (result instanceof Date) {
    return result;
  }
  return result;
}

/**
 * Внутренняя функция для получения даты погашения облигации.
 * @param {string} ticker - Торговый код бумаги.
 * @return {Date | string} - Объект Date или текстовая ошибка.
 */
function fetchMaturityDateInternal(ticker) {
  const url = `https://iss.moex.com/iss/engines/stock/markets/bonds/securities/${encodeURIComponent(
    ticker
  )}.json?iss.meta=off`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    if (
      !data.securities ||
      !data.securities.columns ||
      !data.securities.data ||
      data.securities.data.length === 0
    ) {
      return `Тикер не найден`;
    }

    const columns = data.securities.columns;
    const row = data.securities.data[0];

    const matDateIndex = columns.indexOf('MATDATE');

    if (matDateIndex === -1) {
      return 'Поле MATDATE отсутствует';
    }

    const matDateStr = row[matDateIndex];

    if (!matDateStr || matDateStr === '0000-00-00') {
      return 'Дата погашения не определена';
    }

    return new Date(matDateStr);
  } catch (e) {
    return 'Ошибка скрипта';
  }
}

/**
 * Дополнительная функция для получения купона через bondization.
 * Ищет следующий купон, если его значение null - берет предыдущий известный.
 */
function fetchCouponFromBondization(ticker) {
  const url = `https://iss.moex.com/iss/securities/${encodeURIComponent(ticker)}/bondization.json?iss.meta=off&limit=unlimited`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API (bondization): ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    if (!data.coupons || !data.coupons.data || !data.coupons.columns) {
      return 'Нет данных о купонах (bondization)';
    }

    const columns = data.coupons.columns;
    const rows = data.coupons.data;

    const dateIdx = columns.indexOf('coupondate');
    const valueIdx = columns.indexOf('value');
    const valueRubIdx = columns.indexOf('value_rub');

    if (dateIdx === -1) return 'Нет даты купона в данных';

    // Преобразуем массив в объекты для удобства и сортируем по дате
    const coupons = rows
      .map((r) => {
        // Безопасное получение значения: проверяем индекс и null
        let val = null;
        if (valueRubIdx !== -1 && r[valueRubIdx] !== null) {
          val = r[valueRubIdx];
        } else if (valueIdx !== -1 && r[valueIdx] !== null) {
          val = r[valueIdx];
        }

        return {
          date: new Date(r[dateIdx]),
          value: val !== null ? parseFloat(val) : null,
        };
      })
      .sort((a, b) => a.date - b.date);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let lastKnownValue = null;
    let nextCoupon = null;

    for (let i = 0; i < coupons.length; i++) {
      const c = coupons[i];

      // Если дата купона в будущем (или сегодня)
      if (c.date >= today) {
        nextCoupon = c;
        // Если у следующего купона есть значение и оно не 0, возвращаем его
        if (c.value !== null && !isNaN(c.value) && c.value !== 0) {
          return c.value;
        }
        // Если значения нет или оно 0, прерываем цикл, чтобы вернуть lastKnownValue
        break;
      }

      // Запоминаем последнее известное значение из прошлых купонов
      // Для прошлых купонов 0 может быть валидным значением (хотя редко), но пока оставим как есть
      // Если мы хотим быть строгими к флоатерам, можно тоже игнорировать 0, но это рискованно для других типов.
      // Однако, если мы ищем "предыдущий известный", то 0 вряд ли полезен.
      // Давайте игнорировать 0 и здесь для надежности.
      if (c.value !== null && !isNaN(c.value) && c.value !== 0) {
        lastKnownValue = c.value;
      }
    }

    if (lastKnownValue !== null) {
      return lastKnownValue;
    }

    return 'Купон не определен';
  } catch (e) {
    return 'Ошибка bondization: ' + e.message;
  }
}

/**
 * Кастомная функция для ячейки. Возвращает БЛИЖАЙШУЮ ДАТУ (Put/Call опцион или амортизация).
 * Пытается установить примечание к ячейке с деталями (работает только при запуске из скрипта, не как UDF).
 * @param {string} ticker ISIN или Торговый код облигации.
 * @return {Date | string} Ближайшая дата или текст ошибки.
 * @customfunction
 */
function GET_NEAREST_OPTION_DATE(ticker) {
  if (!ticker) return 'Укажите тикер';

  // Получаем данные
  const result = fetchBondOptionDatesInternal(ticker);

  if (typeof result === 'string') {
    return result; // Возвращаем ошибку
  }

  // Пытаемся установить примечание (Note).
  // ВНИМАНИЕ: В контексте простой формулы (Custom Function) это обычно НЕ работает или запрещено.
  // Но если функция вызывается скриптом, это сработает.
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    // Проверка, не находимся ли мы в контексте ограничения
    if (sheet) {
      const range = SpreadsheetApp.getActiveRange();
      if (range) {
        range.setNote(result.note);
      }
    }
  } catch (e) {
    // Игнорируем ошибку установки примечания, так как в UDF это часто невозможно
    // console.log('Не удалось установить примечание: ' + e.message);
  }

  return result.date;
}

/**
 * Внутренняя функция для получения дат опционов и амортизаций.
 * @param {string} ticker - ISIN или код бумаги.
 * @return {Object | string} - Объект { date: Date, note: string } или строка ошибки.
 */
function fetchBondOptionDatesInternal(ticker) {
  const url = `https://iss.moex.com/iss/securities/${encodeURIComponent(ticker)}/bondization.json?iss.meta=off&limit=unlimited`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return `Ошибка API: ${response.getResponseCode()}`;
    }
    const data = JSON.parse(response.getContentText());

    // Проверяем наличие необходимых блоков данных
    // Обычно это 'amortizations' и 'offers'
    const amortizations = data.amortizations;
    const offers = data.offers;

    // Массив всех найденных будущих событий
    let events = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // 1. Обработка амортизаций
    if (amortizations && amortizations.columns && amortizations.data) {
      const cols = amortizations.columns;
      const dateIdx = cols.indexOf('amortdate');
      // const valIdx = cols.indexOf('value'); // процент или сумма, для инфо (не используется)

      if (dateIdx !== -1) {
        amortizations.data.forEach((row) => {
          const dStr = row[dateIdx];
          if (dStr) {
            const d = new Date(dStr);
            if (d >= today) {
              events.push({
                date: d,
                type: 'Амортизация',
                description: `Амортизация: ${dStr}`,
              });
            }
          }
        });
      }
    }

    // 2. Обработка оферт (Put/Call)
    // В MOEX 'offers' обычно содержит оферты (Put).
    // Call-опционы могут быть там же или в отдельной таблице, но чаще всего 'offers' - это Put/Call.
    // Нужно смотреть поле 'offertype' или подобное, если оно есть.
    if (offers && offers.columns && offers.data) {
      const cols = offers.columns;
      const dateIdx = cols.indexOf('offerdate');
      const typeIdx = cols.indexOf('offertype'); // Может отсутствовать

      if (dateIdx !== -1) {
        offers.data.forEach((row) => {
          const dStr = row[dateIdx];
          if (dStr) {
            const d = new Date(dStr);
            if (d >= today) {
              let type = 'Оферта';
              // Пытаемся уточнить тип, если есть поле offertype
              if (typeIdx !== -1 && row[typeIdx]) {
                type = row[typeIdx]; // Например "Put", "Call" или код
              }

              // Если тип не определен явно, считаем Put (стандартная оферта)
              // Можно добавить логику: если это Call, то помечаем как Call

              events.push({
                date: d,
                type: type, // 'Оферта' (Put) или 'Call'
                description: `${type}: ${dStr}`,
              });
            }
          }
        });
      }
    }

    // Если событий нет
    if (events.length === 0) {
      // Можно вернуть дату погашения как fallback, или сообщение
      // Пользователь просил "выбрать самую ближайшую". Если нет опционов/амортизаций, возможно стоит вернуть MATDATE?
      // Но функция называется GET_NEAREST_OPTION_DATE.
      // Давайте вернем null или сообщение, чтобы не путать с погашением.
      // Или проверим MATDATE?
      // Логичнее вернуть "Нет оферт/аморт."
      return 'Нет оферт/аморт.';
    }

    // Сортируем события по дате
    events.sort((a, b) => a.date - b.date);

    // Ближайшее событие
    const nearest = events[0];

    // Формируем текст примечания со всеми датами
    // Убираем дубликаты описаний, если вдруг
    const uniqueDescriptions = [...new Set(events.map((e) => e.description))];
    const noteText = uniqueDescriptions.join('\n') + `\n\nБлижайшая: ${nearest.description}`;

    return {
      date: nearest.date,
      note: noteText,
    };
  } catch (e) {
    return 'Ошибка скрипта: ' + e.message;
  }
}
