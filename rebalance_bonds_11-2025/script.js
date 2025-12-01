/**
 * Константа: задержка между запросами при массовом обновлении (в миллисекундах)
 */
const DELAY_MS = 400; // 0.4 секунды

/**
 * При открытии документа создает в меню пункт "MOEX".
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MOEX")
    .addItem("Обновить все цены (с задержкой)", "forceRecalculatePrices")
    .addToUi();
}

/**
 * Кастомная функция для ячейки. Возвращает цену облигации по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {number | string} Последняя цена сделки или текстовое описание ошибки.
 * @customfunction
 */
function GET_MOEX_PRICE(ticker) {
  if (!ticker || ticker.trim() === "") {
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
 * Улучшенный обработчик: находит все ячейки с функцией GET_MOEX_PRICE и обновляет их по очереди.
 */
function forceRecalculatePrices() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const targetFunctionName = "GET_MOEX_PRICE";

  const dataRange = sheet.getDataRange();
  const allFormulas = dataRange.getFormulas();
  const targetCells = [];

  for (let i = 0; i < allFormulas.length; i++) {
    for (let j = 0; j < allFormulas[i].length; j++) {
      if (
        allFormulas[i][j] &&
        allFormulas[i][j].toUpperCase().includes(targetFunctionName)
      ) {
        targetCells.push(sheet.getRange(i + 1, j + 1));
      }
    }
  }

  if (targetCells.length === 0) {
    ui.alert(`На листе не найдено ячеек с функцией =${targetFunctionName}().`);
    return;
  }

  ui.alert(
    `Найдено ${targetCells.length} ячеек. Начинаю обновление...`,
    ui.ButtonSet.OK
  );
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Обновление ${targetCells.length} ячеек...`
  );

  targetCells.forEach((cell) => {
    const originalFormula = cell.getFormula();
    cell.clearContent();
    SpreadsheetApp.flush();
    cell.setFormula(originalFormula);
    Utilities.sleep(DELAY_MS);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Обновление цен завершено!",
    "Готово",
    5
  );
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
      marketdataRow[marketdataColumns.indexOf("LAST")] ??
      marketdataRow[marketdataColumns.indexOf("CLOSEPRICE")] ??
      securitiesRow[securitiesColumns.indexOf("PREVLEGALCLOSEPRICE")] ??
      securitiesRow[securitiesColumns.indexOf("PREVPRICE")];

    if (price === null || typeof price === "undefined") {
      return "Цена не найдена"; // Если ни одного значения не нашлось
    }

    return parseFloat(price);
  } catch (e) {
    return "Ошибка скрипта";
  }
}

/**
 * Кастомная функция для ячейки. Возвращает ДАТУ СЛЕДУЮЩЕГО КУПОНА по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {Date | string} Дата следующего купона или текстовое описание ошибки.
 * @customfunction
 */
function GET_NEXT_COUPON(ticker) {
  if (!ticker || ticker.trim() === "") {
    return null;
  }

  // Используем кэш, чтобы не запрашивать одни и те же данные слишком часто
  const cache = CacheService.getScriptCache();
  const cacheKey = ticker + "_coupon"; // Уникальный ключ для кэша купонов
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    // Если дата в кэше, преобразуем ее обратно в объект Date
    const cachedValue = JSON.parse(cached);
    return cachedValue ? new Date(cachedValue) : "Нет данных";
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

    const nextCouponIndex = columns.indexOf("NEXTCOUPON");

    if (nextCouponIndex === -1) {
      return "Поле NEXTCOUPON отсутствует";
    }

    const couponDateStr = row[nextCouponIndex];

    // Проверяем, есть ли дата купона (может не быть у бумаг в обращении или погашенных)
    if (!couponDateStr || couponDateStr === "0000-00-00") {
      return "Нет предстоящих купонов";
    }

    // Возвращаем как объект Date, чтобы Google Sheets правильно понял формат
    return new Date(couponDateStr);
  } catch (e) {
    return "Ошибка скрипта";
  }
}

/**
 * Кастомная функция для ячейки. Возвращает НАИМЕНОВАНИЕ ОБЛИГАЦИИ по тикеру.
 * @param {string} ticker Торговый код облигации (например, "ОФЗ 26227").
 * @return {string} Наименование облигации или текст ошибки.
 * @customfunction
 */
function GET_MOEX_NAME(ticker) {
  if (!ticker || ticker.trim() === "") {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = ticker + "_name";
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
    const secNameIndex = columns.indexOf("SECNAME");
    const shortNameIndex = columns.indexOf("SHORTNAME");

    let name = null;
    if (secNameIndex !== -1) {
      name = row[secNameIndex];
    }
    if (!name && shortNameIndex !== -1) {
      name = row[shortNameIndex];
    }

    if (!name) {
      return "Наименование не найдено";
    }

    return name;
  } catch (e) {
    return "Ошибка скрипта";
  }
}
