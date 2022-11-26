const reqMode = "chain"; 
const textRuApiKey = "63d2b101af327992f6b74937bdc50d5d";

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("Доступ")
        .addItem("Фильтровать", "filterAccess")
        .addToUi();
}

function onOpenClient() {
    SpreadsheetApp.getUi()
      .createMenu("Wildberries")
      .addItem("Целевое ядро (база и хвосты)", "scriptads.findOneAndTwo")
      .addSeparator()
      .addItem("Заполнить доп. поля", "scriptads.fillWBAttributesComplex")
      //.addItem("Заполнить доп. поля в ТЗ", "scriptads.loadWBAttributesList")
      .addItem("Заполнить словарь, запросы, подкатегории", "scriptads.fillDictQueriesExpandsWB")
      //.addItem("Заполнить словарь, запросы, подкатегории (по id предмета)", "scriptads.fillDictQueriesExpandsWBBySubjId")
      //.addItem("Заполнить доп.поля по предмету", "scriptads.fillAdditionalAttributesSubject")
      //.addItem("Заполнить доп.поля по категории", "scriptads.fillAdditionalAttributesCategory")
      //.addItem("Заполнить значения атрибутов в ТЗ", "scriptads.fillAttrValuesForWB")
      .addItem("Комментировать словарь", "scriptads.commentDictWBComplex")
      //.addItem("Комментировать словарь", "scriptads.addDictionaryCommentsWB")
      //.addItem("Комментировать словарь (описание)", "scriptads.commentDescrExcludedDictWords")
      .addSeparator()
      .addItem("Заполнить план выкупов", "scriptads.calcRansomComplex")
      .addSeparator()
      .addItem("Выделить дубли", "scriptads.markRepeatingsInTask")
      .addItem("Проверка вхождений", "scriptads.markTaskExcludedDictWords")
      .addItem("Выделение вхождений в ТЗ", "scriptads.createBold")
      //.addItem("Заполнить план выкупов", "scriptads.calcProcessing")
      //.addItem("Заполнить план выкупов (по id предмета)", "scriptads.calcProcessingBySubjId")
      //.addItem("Заполнить запросы в выкупах", "scriptads.fillRansomQueries")
      .addSeparator()
      .addItem("Отправить текст на проверку", "scriptads.sendToTextRuWb")
      .addItem("Запросить уникальность", "scriptads.getTextUniqueResultsWb")
      .addToUi();
    SpreadsheetApp.getUi()
      .createMenu("Ozon")
      .addItem("Целевое ядро (база и хвосты)", "scriptads.findOneAndTwo")
      .addSeparator()
      .addItem("Заполнить список атрибутов и значений", "scriptads.fillAttributesAndValuesOzon")
      .addItem("Заполнить словарь, запросы", "scriptads.fillDictAndQueriesOzon")
      .addItem("Комментировать словарь", "scriptads.commentDictOzonComplex")
      //.addItem("Комментировать словарь", "scriptads.addDictionaryCommentsOzon")
      //.addItem("Комментировать словарь по описанию", "scriptads.addDictionaryCommentsOzonDescr")
      .addItem("Скрыть ненужные атрибуты", "scriptads.clearOzonFalseAttributes")
      .addSeparator()
      .addItem("Отправить текст на проверку", "scriptads.sendToTextRuOzon")
      .addItem("Запросить уникальность", "scriptads.getTextUniqueResultsOzon")
      .addToUi();
  }
  
////////////// TextRu

function sendToTextRuWb() {
  sendToTextRu("wb");
}

function sendToTextRuOzon() {
  sendToTextRu("oz");
}

function getTextUniqueResultsWb() {
  getTextUniqueResults("wb");
}

function getTextUniqueResultsOzon() {
  getTextUniqueResults("oz");
}

//////////////
function fillWBAttributesComplex() {

    loadWBAttributesList();
  
    try {
      fillAdditionalAttributesSubject()
    } catch (e) {
      if (e.message == 'Не указан предмет WB')
        fillAdditionalAttributesCategory();
    }
  
    fillAttrValuesForWB();
  }
  //////////////
  function commentDictWBComplex() {
  
    addDictionaryCommentsWB();
  
    commentDescrExcludedDictWords();
  }
  
  function commentDictOzonComplex() {
  
    addDictionaryCommentsOzon();
  
    addDictionaryCommentsOzonDescr();
  }
  //////////////

  function calcRansomComplex() {
    calcComplex();
  
    fillRansomQueries(1);
  }
  
  
  function calcComplex() {
    const skuCell = "A4"; // адрес ячейки sku
    const calcSheetName = "Калькулятор"; // название вкладки
  
    const sh = SpreadsheetApp.getActive().getSheetByName(calcSheetName);
  
    if (!sh) throw `Не найден лист ${calcSheetName}`;
  
    let sku = sh.getRange(skuCell).getValue();
  
    if (!sku || !/^\d+$/.test(sku)) {
      calcProcessingBySubjId()
    } else {
      calcProcessing();
    }
  }
  //////////////
  
  function addDictionaryCommentsWB() {
    addDictionaryComments("wb");
  }
  
  function addDictionaryCommentsOzon() {
    addDictionaryComments("ozon");
  }
  
  function fillDictAndQueriesOzon() {
    try {
      let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ТЗ OZON");
  
      let query = sh.getRange("D1").getValue();
      if (!query) throw "Отсутствует запрос в D1";
  
      let skuList = getSkuListBySearch(query);
      let expandingData = expandingRequests(skuList);
  
      fillDictionary(
        expandingData.result.sort((a, b) => b.keys_count_sum - a.keys_count_sum),
        "ozon"
      );
      fillQueries(expandingData.words, "ozon");
    } catch (err) {
      SpreadsheetApp.getUi().alert(err);
    }
  }
  
  function updateDictQueriesColumns() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
  
    let dictSheet = ss.getSheetByName("Словарь");
    if (dictSheet.getRange("F1").getValue().toString().trim() == "Комментарии") {
      dictSheet.insertColumnAfter(4);
      dictSheet.getRange("E1").setValue("Сумм Частота Ozon");
      dictSheet.getRange("F1").setValue("Сумм Частота WB");
    }
  
    let queriesSheet = ss.getSheetByName("Запросы");
    if (
      queriesSheet.getRange("C1").getValue().toString().trim() == "Частотность"
    ) {
      queriesSheet.insertColumnBefore(3);
      queriesSheet.getRange("C1").setValue("Частотность OZON");
      queriesSheet.getRange("D1").setValue("Частотность WB");
    }
  }

  //-------------------------------------Text Ru Api---------------------------------------------------

  function sendToTextRu(mp = "wb") {
    const textCell = mp === "wb" ? "B3" : "C3";
    const resultCell = mp === "wb" ? "H2" : "D2";
    const sheetName = mp === "wb" ? "ТЗ WB" : "ТЗ OZON";
  
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(sheetName);
    let text = sh.getRange(textCell).getValue();
  
    let data = {
      text: text,
      userkey: textRuApiKey
    }
  
    let options = {
      method: "post",
      muteHttpExceptions: true,
      payload: data,
    };
  
    let response = UrlFetchApp.fetch('https://api.text.ru/post', options);
  
    let responseText = response.getContentText();
    let json = JSON.parse(responseText);
  
    if (json.error_desc) {
      SpreadsheetApp.getUi().alert(json.error_desc);
      return;
    }
  
    sh.getRange(resultCell).setValue(json.text_uid);
  
    return json;
  }
  
  function getTextUniqueResults(mp = "wb") {
    const resulCell = mp === "wb" ? "H2" : "D2";
    const sheetName = mp === "wb" ?"ТЗ WB" : "ТЗ OZON";
    
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(sheetName);
  
    let uId = sh.getRange(resulCell).getValue();
  
    if(!uId) {
      SpreadsheetApp.getUi().alert(`Отсутствует id отправленного текста в ячейке ${resulCell}`);
    }
  
    let data = {
      uid: uId,
      userkey: textRuApiKey,
      jsonvisible: "detail"
    };
  
    let options = {
      method: "post",
      muteHttpExceptions: true,
      payload: data,
    };
  
    let response = UrlFetchApp.fetch('https://api.text.ru/post', options);
  
    let responseText = response.getContentText();
    let json = JSON.parse(responseText);
  
    Logger.log(JSON.stringify(json));
    if (json.error_desc) {
      SpreadsheetApp.getUi().alert(json.error_desc);
      return;
    }
  
    if (json.text_unique) {
      sh.getRange(resulCell).setValue(`Уникальность описания: ${json.text_unique}%`);
    } else {
      SpreadsheetApp.getUi().alert("Ошибка. Отсутствуют данные результата");
    }
  }

  
  //--------------------------------Заполнение с mpstats------------------------------------------------
  
  // Заполнить словарь, запросы, подкатегории
  function fillDictQueriesExpandsWB() {
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ТЗ WB");
      let c1 = sh.getRange("C1").getValue().toString();
      let sku;
      //debugger
  
      if (/^\d+$/.test(c1)) {
        sku = c1;
      } else {
        let d1 = sh.getRange("D1").getValue().toString();
        if (/^\d+$/.test(d1)) {
          sku = d1;
        } else {
          fillDictQueriesExpandsWBBySubjId();
          return;
          //throw "Нет доступного SKU WB для определения предмета";
        }
      }
  
      Logger.log("Start heroku req");
      let subjectId = herokuGetMpstatsData(sku).subjId;
  
      Logger.log(`Recieved heroku res ${subjectId}`);
  
      let dateFrom = sh.getRange("F1").getValue();
  
      if (dateFrom) {
        if (dateFrom.toString().toLowerCase() == "date") dateFrom = null;
        else dateFrom = makeDate(dateFrom);
      }
  
      Logger.log("Get sku list");
      let skuList = getSkuList(subjectId, dateFrom);
  
      Logger.log("expandingRequests");
      let expandingData = expandingRequests(skuList, dateFrom);
  
      Logger.log("getSubCategoriesList");
      let subCategories = getSubCategoriesList(skuList);
  
  
      fillDictionary(
        expandingData.result.sort(
          (a, b) => b.keys_wb_count_sum - a.keys_wb_count_sum
        )
      );
  
      fillQueries(expandingData.words.sort((a, b) => b.wbcount - a.wbcount));
  
      fillSubcategories(subCategories);
  
    } catch (err) {
      SpreadsheetApp.getUi().alert(err);
    }
  }
  
  // Заполнить словарь, запросы, подкатегории
  function fillDictQueriesExpandsWBBySubjId() {
    let subjIdCell = "G1";
  
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ТЗ WB");
      let subjectId = sh.getRange(subjIdCell).getValue().toString();
  
      if (!subjectId || !/^\d+$/.test(subjectId))
        throw `Не ведён доступный id предмета в ячейке ${subjectId}`;
  
      let dateFrom = sh.getRange("F1").getValue();
  
      if (dateFrom) {
        if (dateFrom.toString().toLowerCase() == "date") dateFrom = null;
        else dateFrom = makeDate(dateFrom);
      }
  
      let skuList = getSkuList(subjectId, dateFrom);
  
      let expandingData = expandingRequests(skuList, dateFrom);
      let subCategories = getSubCategoriesList(skuList);
  
      fillDictionary(
        expandingData.result.sort(
          (a, b) => b.keys_wb_count_sum - a.keys_wb_count_sum
        )
      );
      fillQueries(expandingData.words.sort((a, b) => b.wbcount - a.wbcount));
      fillSubcategories(subCategories);
    } catch (err) {
      SpreadsheetApp.getUi().alert(err);
    }
  }
  
  function makeDate(string) {
    if (Object.prototype.toString.call(string) === "[object Date]") {
      return string;
    }
  
    let splitted = string.toString().split(/[.,]/);
  
    if (splitted.length < 2)
      throw "Дата отсчета для сбора данных введена некорректно";
  
    let day = +splitted[0] + 1;
    let month = splitted[1];
    let year = splitted.length > 2 ? splitted[2] : new Date().getFullYear();
  
    if (year.length == 2) year = `20${year}`;
  
    return new Date(year, month - 1, day);
  }
  
  function testDebug() {
    let skuList = getSkuList(1440);
    let expandingData = expandingRequests(skuList);
    let subCategories = getSubCategoriesList(skuList);
  
    fillDictionary(
      expandingData.result.sort((a, b) => b.keys_count_sum - a.keys_count_sum)
    );
    fillQueries(expandingData.words);
    fillSubcategories(subCategories);
  }
  
  function fillAdditionalAttributesSubject() {
    let subject = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .getRange("A1")
      .getValue();
    if (!subject) throw "Не указан предмет WB";
    fillAttributes(subject, "Доп.Поля");
  }
  
  function fillAdditionalAttributesCategory() {
    let category = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .getRange("B1")
      .getValue();
    if (!category) throw "Не указана категория WB";
    fillAttributes(category, "Доп.Поля", false);
  }
  
  function fillAttrValuesForWB() {
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName("ТЗ WB");
    let range = sh.getRange("A4:B");
    let values = range.getValues();
  
    let shExtAttrs = ss.getSheetByName("Доп.Поля");
    if (!shExtAttrs) throw "Сначала заполните Доп.Поля";
    let attrValues = shExtAttrs
      .getRange(1, 1, shExtAttrs.getLastRow(), shExtAttrs.getLastColumn())
      .getValues();
  
    values.forEach((v) => {
      if (!v[0]) return;
      let findRow = attrValues.find((el) => el[1] === v[0].toString());
      if (findRow)
        v[1] = `(Макс. полей ${findRow[4]}); ${findRow
          .filter((v) => v)
          .slice(5)
          .join("; ")}`;
    });
  
    range.setValues(values);
  }
  
  function fillDictionary(data, market = "wb") {
    updateDictQueriesColumns();
  
    let values = data.map((v) => [
      v.word,
      v.words.join(","),
      v.count,
      v.keys_count_sum,
      v.keys_wb_count_sum,
    ]);
  
    fillTableData("Словарь", "B", "F", 2, values);
  
    let dictSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Словарь");
  
    dictSheet.setTabColor(market === "wb" ? "#9400d3" : "#6495ed");
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .setTabColor(market === "wb" ? "#9400d3" : null);
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ OZON")
      .setTabColor(market === "wb" ? null : "#6495ed");
  
    dictSheet.hideColumns(market == "wb" ? 5 : 6);
    dictSheet.showColumns(market == "wb" ? 6 : 5);
  
    clearRange("Словарь", "G", "G", 2);
    clearRange("Словарь", "A", "A", 2);
  }
  
  function fillQueries(data, market = "wb") {
    updateDictQueriesColumns();
  
    let values = data.map((v) => [v.word, v.count, v.wbcount]);
    fillTableData("Запросы", "B", "D", 2, values);
  
    let queriesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Запросы");
  
    queriesSheet.setTabColor(market === "wb" ? "#9400d3" : "#6495ed");
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .setTabColor(market === "wb" ? "#9400d3" : null);
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ OZON")
      .setTabColor(market === "wb" ? null : "#6495ed");
  
    queriesSheet.hideColumns(market == "wb" ? 3 : 4);
    queriesSheet.showColumns(market == "wb" ? 4 : 3);
  }
  
  function fillSubcategories(data) {

    let values = data
      .map((v) => [v.category])
      .filter((v, i, ar) => {
  
        if (v[0].toString().startsWith("Акции"))
          return false;
  
        if (i < ar.length - 1 && ar[i + 1][0].replace(/\/[^\/]*$/, '').includes(v))
          return false;
  
        return true;
      }).filter((v, i, ar) => {
      if (i < ar.length - 1) {
        let splitCurCategories = v[0].trim().split('/');
        let curWords = splitCurCategories[splitCurCategories.length - 1].split(' ');
  
        let splitNextCategories = ar[i + 1][0].trim().split('/');
        let nextWords = splitNextCategories[splitNextCategories.length - 1].split(' ');
  
        let subSame = false;
  
        let curPath = v[0].replace(/\/[^\/]*$/, '');
        let nextPath = ar[i + 1][0].replace(/\/[^\/]*$/, '');
  
        if (curPath == nextPath)
          subSame = true;
  
        let isOk = true;
  
        if (subSame
          && curWords.length == nextWords.length
          && curWords.every((v, i) => stemmer(v).startsWith(stemmer(nextWords[i])))) {
          isOk = false;
        }
  
        return isOk;
      }
  
      return true;
    });
  
    values.forEach(v => Logger.log(v[0]));
  
    if (!values.length) {
      let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Подкатегории");
      sh.getRange("B2").setValue("Нет категорий");
    } else {
      fillTableData("Подкатегории", "B", "B", 2, values);
    }
  }
  
  function fillTableData(sheetName, fCln, lCln, startRow, values) {
    let columns = `${fCln}${startRow}:${lCln}`;
  
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(sheetName);
  
    if (sh == null) {
      sh = ss.insertSheet(sheetName);
    }
  
    let rows = sh.getMaxRows();
    let rowsDiff = values.length - rows - (startRow - 1);
  
    if (rowsDiff > 0) {
      sh.insertRows(rows, rowsDiff);
    }
    sh.getRange(columns).clearContent();
    sh.getRange(`${columns}${values.length + (startRow - 1)}`).setValues(values);
  }
  
  function clearRange(sheetName, fCln, lCln, startRow) {
    let columns = `${fCln}${startRow}:${lCln}`;
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(sheetName);
  
    if (sh == null) throw `Листа ${sheetName} не существует`;
  
    sh.getRange(columns).clearContent();
  }
  
  function setColor(sheetName, fCln, lCln, startRow, color) {
    let columns = `${fCln}${startRow}:${lCln}`;
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(sheetName);
  
    if (sh == null) throw `Листа ${sheetName} не существует`;
  
    sh.getRange(columns).setFontColor(color);
  }
  
  function getSkuList(subjectId, dateFromSheet = null, lastDays = 30) {
    let dateTo = new Date();
    let dateFrom = new Date();
  
    if (dateFromSheet) {
      dateFrom = dateFromSheet;
      dateTo = new Date(dateFrom);
      dateTo.setDate(dateTo.getDate() + 30);
    } else {
      dateFrom.setDate(dateFrom.getDate() - lastDays);
      dateTo.setDate(dateTo.getDate() - 1);
    }
  
    dateTo = dateTo.toISOString().split("T")[0];
    dateFrom = dateFrom.toISOString().split("T")[0];
  
    /* example **
    "filterModel": {
      "category": {
        "filterType":"text",
        "operator":"OR",
        "condition1": {
          "filterType":"text",
          "type":"contains",
          "filter":"дет"
          },
        "condition2": {
          "filterType":"text",
          "type":"contains",
          "filter":"дев"
        }
      }
    }
   */
  
    let data = {
      dataFilter: {},
      filterModel: constructFilter(),
      sortModel: [{ colId: "revenue", sort: "desc", sortIndex: 0 }],
      fields: ["id", "category", "revenue"],
    };
  
    let response = mpStatsRequest(
      `wb/get/subject?d1=${dateFrom}&d2=${dateTo}&path=${subjectId}&type=json`,
      data
    );
  
    if (!response.data || !response.data.length) throw "Нет данных";
  
    return response.data.map((v) => v.id);
  }
  
  function expandingRequests(skuList, dateFromSheet = null, lastDays = 30) {
    let dateTo = new Date();
    let dateFrom = new Date();
  
    if (dateFromSheet) {
      dateFrom = dateFromSheet;
      dateTo = new Date(dateFrom);
      dateTo.setDate(dateTo.getDate() + 30);
    } else {
      dateFrom.setDate(dateFrom.getDate() - lastDays);
      dateTo.setDate(dateTo.getDate() - 1);
    }
  
    dateTo = dateTo.toISOString().split("T")[0];
    dateFrom = dateFrom.toISOString().split("T")[0];
  
    let data = {
      query: skuList.join(","),
      type: "sku",
      mp: 0,
      similar: false,
      stopWords: [],
      searchFullWord: false,
      d1: dateFrom,
      d2: dateTo,
    };
  
    let response = mpStatsRequest("seo/keywords/expanding", data);
  
    if (response && response.result.length == 0)
      throw "Отсутствуют данные по расширению запросов";
  
    return response;
  }
  
  function getSubCategoriesList(skuList) {
    return mpStatsRequest("seo/tools/wb-sku-checker", { sku: skuList });
  }
  
  function mpStatsRequest(url = "", data = {}) {
    if (reqMode && reqMode === "chain") {
      return mpStatsRequestHeroku(url, data);
    }
  
    return mpStatsRequestDirect(url, data);
  }
  
  function mpStatsRequestDirect(url = "", data = {}) {
    let apiKey = "62e010f690d6f6.72414271c10c7a6b8873f0b809951a51ebf31e69";
  
    let options = {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: {
        "X-Mpstats-TOKEN": apiKey,
      },
      payload: JSON.stringify(data),
    };
  
    let response;
  
    try {
      Logger.log(`Start mpstats req at ${url}`);
  
      response = UrlFetchApp.fetch(`https://mpstats.io/api/${url}`, options);
  
      Logger.log(`Recieved mpstats res`);
  
    } catch (e) {
      Logger.log(`${e.message}`);
  
      if (e.message.includes("Address unavailable"))
        throw "Mpstats Api не ответил, попробуйте повторить запрос";
  
      else throw e;
    }
  
    return JSON.parse(response.getContentText());
  }
  
  function mpStatsRequestHeroku(url = "", data = {}) {
    let options = {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({ url, data }),
    };
  
    Logger.log(`start h-mpstat at ${url}`);
  
    let response, responseText, nTries = 0;
  
    do {
      response = UrlFetchApp.fetch(`https://zmpstat.herokuapp.com/mpstats/api/`, options);
  
      responseText = response.getContentText()
  
      if (responseText.includes('error-pages'))
        Logger.log("error");
  
    } while (responseText.includes('error-pages') && nTries++ < 5)
  
    Logger.log("finish h-mpstat");
  
    return JSON.parse(responseText);
  }
  
  function fillAttributes(param, sheetName, isSubject = true) {
    let partial = isSubject ? "subject?subject" : "category?catname";
    let response = UrlFetchApp.fetch(`http://65.108.9.60:7007/get_stats_by_${partial}=${encodeURI(param)}&numtop=100&minfreq=40&mincats=0`,
      {
        headers: {
          Authorization:
            "Basic bXBzdGF0c19jb25zdWx0aW5nOjkxZjc2MWZjOGRhQWYxN2E3NzhERjc=",
        },
      }
    );
  
    let xlsx = response.getBlob();
  
    var file = {
      title: new Date().getTime(),
    };
    let spreadsheet = Drive.Files.insert(file, xlsx, { convert: true });
  
    let curss = SpreadsheetApp.getActiveSpreadsheet();
  
    let cursh = curss.getSheetByName(sheetName);
    if (cursh) curss.deleteSheet(cursh);
  
    let ss = SpreadsheetApp.openById(spreadsheet.id);
    let sh = ss.getSheetByName("Атрибуты");
  
    sh.copyTo(curss);
    curss.getSheetByName("Атрибуты (копия)").setName(sheetName).hideSheet();
  
    Drive.Files.remove(spreadsheet.id);
  }
  
  function findGender() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName("ТЗ WB");
  
    let values = sh.getRange("A4:B").getValues();
  
    let gender = values.filter((v) => v[0] === "Пол")[0];
  
    return gender ? gender[1].slice(0, 3) : null;
  }
  
  function constructFilter() {
    let gender = findGender();
    return gender
      ? { category: { filterType: "text", type: "contains", filter: gender } }
      : {};
  }
  
  //----------------------------------------------------------------------------------------------------
  
  //------------------------------------WB API----------------------------------------------------------
  function loadWBAttributesList() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = "ТЗ WB";
    let subjectCell = "A1";
    let yPos = 4;
    let filterValues = [
      "Бренд",
      "Описание",
      "Ключевые слова",
      "Комплектация",
      "Состав",
    ];
  
    let dashWB = spreadsheet.getSheetByName("Дашборд WB");
    if (dashWB != null) dashWB.showSheet().setTabColor("#9400d3");
  
    let ozonWB = spreadsheet.getSheetByName("Дашборд OZON");
    if (ozonWB != null) ozonWB.hideSheet();
  
    spreadsheet.getSheetByName("ТЗ WB").showSheet().setTabColor("#9400d3");
    spreadsheet.getSheetByName("ТЗ OZON").hideSheet();
    spreadsheet.getSheetByName("Подкатегории").showSheet();
  
    let offset = yPos - 1;
    let sh = spreadsheet.getSheetByName(sheetName);
  
    let subj = sh.getRange(subjectCell).getValue();
    if (!subj)
      throw `Не введено название предмета в ячейке ${subjectCell} на листе ${sheetName}`;
  
    let response = UrlFetchApp.fetch(
      `https://content-suppliers.wildberries.ru/ns/characteristics-configurator-api/content-configurator/api/v1/config/get/object/translated?name=${encodeURI(
        subj
      )}&lang=ru`
    );
  
    response = JSON.parse(response.getContentText());
  
    if (!response.data)
      throw `Нет данных по введенному в ${subjectCell} предмету`;
  
    let attrs = response.data.addin;
  
    attrs = attrs
      .filter((v) => !filterValues.includes(v.type))
      .map((v) => [v.type]);
  
    var range = sh.getRange(`A${yPos}:A`);
    range.clearContent();
  
    let diffRows = attrs.length - range.getNumRows();
  
    if (diffRows > 0) {
      sh.insertRows(range.getNumRows(), diffRows);
    }
  
    range = sh.getRange(`A${yPos}:A${attrs.length + offset}`).setValues(attrs);
  }
  
  //----------------------------------------------------------------------------------------------------
  
  function findOneAndTwo() {
    findInArr(["1", "2"]);
  }
  
  function findInArr(numbers) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s1 = ss.getSheetByName("Словарь");
    var s2 = ss.getSheetByName("Запросы");
  
    s2.getRange("A2:A").clearContent();
  
    var ranksSrc = s1.getRange("A2:A").getValues();
    var ranksTarg = s2.getRange("A2:A").getValues();
  
    var queries = s2
      .getRange("B2:B")
      .getValues()
      .map((x) => x[0].toString().toUpperCase());
    queries.splice(lastNotEmpty(queries) + 1);
  
    var words = s1
      .getRange("C2:C")
      .getValues()
      .map((x) => x[0].toString().toUpperCase().split(","));
    words.splice(lastNotEmpty(words) + 1);
  
    queries.forEach((value, index) => {
      let findInd;
      let queryWords = value.split(" ");
  
      if (
        queryWords.some((word) => {
          findInd = words.findIndex((x) => x.includes(word));
          return (
            findInd > -1 && numbers.includes(ranksSrc[findInd][0].toString())
          );
        }) &&
        queryWords.every((word) => {
          let find = words.findIndex((x) => x.includes(word));
          return !(find > -1 && !numbers.includes(ranksSrc[find][0].toString()));
        })
      ) {
        ranksTarg[index][0] = ranksSrc[findInd][0];
      }
    });
  
    s2.getRange("A2:A").setValues(ranksTarg);
  }
  
  // индекс последнего не пустого элемента массива
  function lastNotEmpty(array) {
    index = array.length;
    while (!array[--index]);
    return index;
  }
  
  // выделить в ТЗ жирным слова, имеющиеся в словаре с пометками 1, 2
  function createBold() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("ТЗ WB");
    sh.getRange("H4:H").clearContent();
  
    var range = sh.getRange(2, 2, sh.getLastRow() - 1, 3);
    range.setFontColor("black");
    range.setFontWeight("normal");
  
    var vals = range.getValues();
  
    var dict = getWords();
  
    vals.forEach((row, rowInd) =>
      row.forEach((txt, i) => {
        if (Object.prototype.toString.call(txt) === "[object Date]") {
          txt = txt.toLocaleDateString("ru-ru", {
            day: "numeric",
            month: "long",
          });
        }
        let ltxt = txt.toString().toLowerCase();
        let results = [];
  
        dict.forEach((wrd) => {
          results = results.concat(
            Array.from(
              ltxt.matchAll(
                new RegExp(`(?<=^|[^а-яА-ЯЁё\\w])${stemmer(wrd)}[a-яё\\w]*`, "gi")
              )
            )
          );
        });
  
        if (results.length > 0) {
          var richValue = SpreadsheetApp.newRichTextValue();
          var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
          richValue.setText(txt);
  
          results.forEach((r) => {
            richValue.setTextStyle(r.index, r.index + r[0].length, bold);
          });
  
          sh.getRange(rowInd + 2, i + 2).setRichTextValue(richValue.build());
        }
      })
    );
  }
  
  // выделить цветом в словаре слова (1,2), отсутствующие в ТЗ
  function markTaskExcludedDictWords() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("Словарь");
    sh.getRange("B2:B").setFontColor("black");
  
    // диапазон столбец A,B,C со второй строки
    var range = sh.getRange(2, 1, sh.getLastRow() - 1, 3);
    var values = range.getValues();
  
    // список слов в ТЗ
    let taskWords = Array.from(getTaskUniqueWords());
  
    // проход по строкам диапазона
    values.forEach((row, rowInd) => {
      // только при значениях 1 и 2 в колонке А
      if (
        (row[0] == 1 || row[0] == 2) &&
        // проверка на вхождение в какое-либо слово из ТЗ
        !taskWords.some((w) => {
          let taskWord = w.toUpperCase().replace("Ё", "Е");
          let wordForms = row[2]
            .toString()
            .split(/\s*,\s*/)
            .filter((s) => s.length);
          // проверка на вхождение слова из текущей строки столбца B
          let wordsMatch =
            taskWord.indexOf(stemmer(row[1].toString()).toUpperCase()) > -1;
          // проверка на вхождение какого-либо слова из текущей строки столбца C
          let wordFormsMatch = wordForms.some(
            (w) => taskWord.indexOf(stemmer(w.toString()).toUpperCase()) > -1
          );
  
          return wordsMatch || (wordForms.length && wordFormsMatch);
        })
      ) {
        range.getCell(rowInd + 1, 2).setFontColor("red");
      }
    });
  }
  
  // выделить в ТЗ слова с повторяющимися основами
  function markRepeatingsInTask() {
    // От какого количества вхождений выделять
    const minOccurences = 2;
  
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("ТЗ WB");
  
    var range = sh.getRange(2, 2, sh.getLastRow() - 1, 3);
    range.setFontColor("black");
    range.setFontWeight("normal");
  
    var vals = range.getValues();
  
    let occurs = getTaskWordsCounts();
    var ocObj = [];
    let n = 0;
  
    // создание массива объектов {слово, количество, цвет} для нужного количества вхождений
    for (const [key, value] of Object.entries(occurs)) {
      if (value >= minOccurences)
        ocObj.push({
          word: key,
          count: value,
          color: COLORS[n++ % COLORS.length],
        });
    }
  
    ocObj.sort((a, b) => b.count - a.count);
  
    // вывод в колонку информации о повторениях
    let infoRange = sh.getRange("H4:H");
    infoRange.clearContent();
    infoRange.setHorizontalAlignment("left");
  
    // добавить строки, если не хватает для записи
    let rowsCountDiff = sh.getLastRow() - 3 - ocObj.length;
    if (rowsCountDiff < 0) {
      sh.insertRowsAfter(sh.getLastRow(), Math.abs(rowsCountDiff));
    }
  
    ocObj.forEach((obj, ind) => {
      var richValue = SpreadsheetApp.newRichTextValue();
      richValue.setText(`${obj.word} - ${obj.count}`);
      var style = SpreadsheetApp.newTextStyle()
        .setForegroundColor(obj.color)
        .setBold(true)
        .build();
      richValue.setTextStyle(style);
      sh.getRange("H" + (4 + ind)).setRichTextValue(richValue.build());
    });
  
    // выделение в тексте цветом соответствующих слов по количеству вхождений
    vals.forEach((row, rowInd) =>
      row.forEach((txt, i) => {
        let ltxt = txt;
        if (Object.prototype.toString.call(ltxt) === "[object Date]")
          ltxt = ltxt.toLocaleDateString("ru-ru", {
            day: "numeric",
            month: "long",
          });
        ltxt = txt.toString().toUpperCase();
        let results = [];
  
        ocObj.forEach((o) => {
          let result = Array.from(
            ltxt.matchAll(
              new RegExp(
                `(?<=(^|[^а-яА-ЯЁё\\w]))${o.word.toString()}[a-яё\\w]*`,
                "gi"
              )
            )
          );
          results = results.concat(result);
        });
  
        if (results.length > 0) {
          var richValue = SpreadsheetApp.newRichTextValue();
          richValue.setText(txt);
          results.forEach((r) => {
            let colorName = ocObj.filter(
              (o) =>
                r[0].toString().toUpperCase().indexOf(o.word.toString()) === 0
            )[0].color;
            var color = SpreadsheetApp.newTextStyle()
              .setForegroundColor(colorName)
              .setBold(true)
              .build();
            richValue.setTextStyle(r.index, r.index + r[0].length, color);
          });
          sh.getRange(rowInd + 2, i + 2).setRichTextValue(richValue.build());
        }
      })
    );
  }
  
  // Откомментировать в словаре при вхождении словоформ в заголовке/значениях доп.полей
  function addDictionaryComments(market = "wb") {
    updateDictQueriesColumns();
  
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("Словарь");
  
    sh.setTabColor(market === "wb" ? "#9400d3" : "#6495ed");
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .setTabColor(market === "wb" ? "#9400d3" : null);
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ OZON")
      .setTabColor(market === "wb" ? null : "#6495ed");
  
    let range = sh.getRange("A2:G");
    var values = range.getValues();
  
    let titleWords = getTitleUniqeWords(market);
    let attrWords =
      market === "wb" ? getAttrsUniqueWords() : getAttrsUniqueWordsOzon();
  
    let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
  
    values.forEach((v) => {
      let words = v[2].toString().match(pattern);
      if (!words) return;
  
      v[6] = "";
  
      words.forEach((w) => {
        let val = w.toString().toLowerCase();
  
        if (
          titleWords &&
          titleWords.includes(val) &&
          !v[6].includes("Заголовок")
        ) {
          v[6] = v[6] ? `Заголовок, ${v[6]}` : "Заголовок";
          v[0] = "1";
        }
        if (attrWords && attrWords.includes(val) && !v[6].includes("Доп.поле")) {
          v[6] += `${v[6] ? ", " : ""}Доп.поле`;
          v[0] = "1";
        }
      });
    });
    range.setValues(values);
  }
  
  // Откомментировать в словаре при вхождении словоформ в описании Озон
  function addDictionaryCommentsOzonDescr() {
    updateDictQueriesColumns();
  
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("Словарь");
  
    sh.setTabColor("#6495ed");
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ WB")
      .setTabColor(null);
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("ТЗ OZON")
      .setTabColor("#6495ed");
  
    let range = sh.getRange("A2:G");
    var values = range.getValues();
  
    let descrWords = cellUniqueWords("C3", "ТЗ OZON");
  
    let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
  
    values.forEach((v) => {
      let words = v[2].toString().match(pattern);
      if (!words) return;
  
      words.forEach((w) => {
        let val = w.toString().toLowerCase();
  
        if (
          descrWords &&
          descrWords.includes(val) &&
          !v[6].includes("Описание")
        ) {
          v[6] += `${v[6] ? ", " : ""}Описание`;
          v[0] = "1";
        }
      });
    });
    range.setValues(values);
  }
  
  function commentDescrExcludedDictWords() {
    updateDictQueriesColumns();
  
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("Словарь");
  
    let range = sh.getRange("A2:G");
    let values = range.getValues();
    let descrWords = getDescriptionUniqeWords();
  
    //let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
  
    values.forEach((v) => {
      
      if (!v[1]) return;
  
      let word = v[1];
      if (Object.prototype.toString.call(word) === "[object Date]")
        word = word.toLocaleDateString("ru-ru", {
          day: "numeric",
          month: "long",
        });
      word = word.toString();
  
      if (descrWords.some((w) => w.startsWith(stemmer(word)))) {
        if(!v[0]) 
          v[0] = "1";
        if(!v[6]) 
          v[6] = "Описание";
      }
    });
    range.setValues(values);
  }
  
  // Список слов из заголовка
  function getTitleUniqeWords(market = "wb") {
    let cell = `${market === "wb" ? "B" : "C"}2`;
    let sheet = `ТЗ ${market === "wb" ? "WB" : "OZON"}`;
    return cellUniqueWords(cell, sheet);
  }
  
  // Список слов из описания
  function getDescriptionUniqeWords() {
    return cellUniqueWords("B3", "ТЗ WB");
  }
  
  function cellUniqueWords(cell, sheet) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(sheet);
    var val = sh.getRange(cell).getValue();
  
    let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
    return val
      ?.toString()
      .match(pattern)
      ?.filter((v, i, a) => v && a.indexOf(v) === i)
      .map((v) => v.toLocaleLowerCase());
  }
  
  // Список слов из значений доп.полей
  function getAttrsUniqueWords(market = "wb") {
    let sheetName = `ТЗ ${market === "wb" ? "WB" : "OZON"}`;
    let range = market === "wb" ? "B4:D" : "C4:C";
  
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(sheetName);
    let values = sh.getRange(range).getValues();
  
    let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
    let words = [];
    values.forEach((v) => {
      v.forEach((v2) => {
        if (!v2) return;
        let tmp = v2;
        if (Object.prototype.toString.call(tmp) === "[object Date]")
          tmp = tmp.toLocaleDateString("ru-ru", {
            day: "numeric",
            month: "long",
          });
        tmp = tmp.toString().match(pattern);
        if (!tmp) return;
        tmp = tmp?.map((v) => v.toLowerCase());
        words = words.concat(tmp);
      });
    });
  
    return words.filter((v, i, a) => v && a.indexOf(v) === i);
  }
  
  // Список слов из значений доп.полей
  function getAttrsUniqueWordsOzon() {
    let sheetName = `ТЗ OZON`;
    let range = "B4:C";
  
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(sheetName);
    let values = sh.getRange(range).getValues();
  
    let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
    let words = [];
  
    values.forEach((v) => {
      if (!v[0]) return;
      let tmp = v[1].toString().match(pattern);
      if (!tmp) return;
      tmp = tmp?.map((v) => v.toLowerCase());
      words = words.concat(tmp);
    });
  
    return words.filter((v, i, a) => v && a.indexOf(v) === i);
  }
  
  // получить список уникальных слов из задания
  function getTaskUniqueWords() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("ТЗ WB");
  
    var range = sh.getRange(2, 2, sh.getLastRow() - 1, 3);
    var values = range.getValues();
  
    let words = new Set();
  
    values.forEach((row, rowInd) =>
      row.forEach((txt, colInd) => {
        if (Object.prototype.toString.call(txt) === "[object Date]")
          txt = txt.toLocaleDateString("ru-ru", {
            day: "numeric",
            month: "long",
          });
        let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
        let matches = Array.from(txt.toString().matchAll(pattern));
        matches.forEach((m) => words.add(m[0]));
      })
    );
  
    return words;
  }
  
  // получить словарь [слово-вхождения] из задания
  function getTaskWordsCounts() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("ТЗ WB");
  
    var range = sh.getRange(2, 2, sh.getLastRow() - 1, 3);
    var values = range.getValues();
  
    let words = {};
  
    values.forEach((row, rowInd) =>
      row.forEach((txt, colInd) => {
        if (Object.prototype.toString.call(txt) === "[object Date]")
          txt = txt.toLocaleDateString("ru-ru", {
            day: "numeric",
            month: "long",
          });
        let pattern = /[а-яА-ЯёЁ\w]{3,}|\d+/gi;
        let matches = Array.from(txt.toString().matchAll(pattern));
        matches.forEach((m) => {
          let word = stemmer(m[0].toString()).toUpperCase();
          words[word] = words[word] ? words[word] + 1 : 1;
        });
      })
    );
  
    return words;
  }
  
  // получить массив уникальных слов вкладки "Словарь"
  function getWords() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("Словарь");
    var vals = sh
      .getRange(2, 1, sh.getLastRow() - 1, 3)
      .getValues()
      .filter((r) => r[0] == 1 || r[0] == 2);
    var words = vals.map((r) => r[1]);
    words = words.filter((r) => r.length > 1);
    var texts = vals.map((r) => r[2]);
    texts.forEach(function (txt) {
      if (Object.prototype.toString.call(txt) === "[object Date]")
        txt = txt.toLocaleDateString("ru-ru", { day: "numeric", month: "long" });
      txt
        .toString()
        .split(",")
        .forEach((wrd) => {
          if (wrd.length > 1) {
            words.push(wrd);
          }
        });
    });
    return words.filter((value, index, self) => self.indexOf(value) === index);
  }
  
  // обработка окончаний/суффиксов слова
  var stemmer = (function () {
    var DICT = {
      RVRE: /^(.*?[аеиоуыэюя])(.*)$/i,
      PERFECTIVEGROUND_1: /([ая])(в|вши|вшись)$/gi,
      PERFECTIVEGROUND_2: /(ив|ивши|ившись|ыв|ывши|ывшись)$/i,
      REFLEXIVE: /(с[яь])$/i,
      ADJECTIVE:
        /(ее|ие|ые|ое|ими|ыми|ей|ий|ый|ой|ем|им|ым|ом|его|ого|ему|ому|их|ых|ую|юю|ая|яя|ою|ею)$/i,
      PARTICIPLE_1: /([ая])(ем|нн|вш|ющ|щ)$/gi,
      PARTICIPLE_2: /(ивш|ывш|ующ)$/i,
      VERB_1: /([ая])(ла|на|ете|йте|ли|й|л|ем|н|ло|но|ет|ют|ны|ть|ешь|нно)$/gi,
      VERB_2:
        /(ила|ыла|ена|ейте|уйте|ите|или|ыли|ей|уй|ил|ыл|им|ым|ен|ило|ыло|ено|ят|ует|уют|ит|ыт|ены|ить|ыть|ишь|ую|ю)$/i,
      NOUN: /(а|ев|ов|ие|ье|е|иями|ями|ами|еи|ии|и|ией|ей|ой|ий|й|иям|ям|ием|ем|ам|ом|о|у|ах|иях|ях|ы|ь|ию|ью|ю|ия|ья|я)$/i,
      DERIVATIONAL: /.*[^аеиоуыэюя]+[аеиоуыэюя].*ость?$/i,
      DER: /ость?$/i,
      SUPERLATIVE: /(ейше|ейш)$/i,
      I: /и$/i,
      P: /ь$/i,
      NN: /нн$/i,
    };
  
    return function stemmer(word) {
      word = word.replace(/ё/gi, "е");
      var wParts = word.match(DICT.RVRE);
      if (!wParts) {
        return word;
      }
      var start = wParts[1];
      var rv = wParts[2];
      var temp = rv.replace(DICT.PERFECTIVEGROUND_2, "");
      if (temp == rv) {
        temp = rv.replace(DICT.PERFECTIVEGROUND_1, "$1");
      }
      if (temp == rv) {
        rv = rv.replace(DICT.REFLEXIVE, "");
        temp = rv.replace(DICT.ADJECTIVE, "");
        if (temp != rv) {
          rv = temp;
          temp = rv.replace(DICT.PARTICIPLE_2, "");
          if (temp == rv) {
            rv = rv.replace(DICT.PARTICIPLE_1, "$1");
          }
        } else {
          temp = rv.replace(DICT.VERB_2, "");
          if (temp == rv) {
            temp = rv.replace(DICT.VERB_1, "$1");
          }
          if (temp == rv) {
            rv = rv.replace(DICT.NOUN, "");
          } else {
            rv = temp;
          }
        }
      } else {
        rv = temp;
      }
      rv = rv.replace(DICT.I, "");
      if (rv.match(DICT.DERIVATIONAL)) {
        rv = rv.replace(DICT.DER, "");
      }
      temp = rv.replace(DICT.P, "");
      if (temp == rv) {
        rv = rv.replace(DICT.SUPERLATIVE, "");
        rv = rv.replace(DICT.NN, "н");
      } else {
        rv = temp;
      }
      return start + rv;
    };
  })();
  
  // Генерация случайного целого числа
  function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1) + min);
  }
  
  const COLORS = [
    "gray",
    "maroon",
    "red",
    "purple",
    "fuchsia",
    "green",
    "Navy",
    "olive",
    "blue",
    "teal",
    "blueviolet",
    "brown",
    "chocolate",
    "coral",
    "cornflowerblue",
    "crimson",
    "darkviolet",
    "deeppink",
    "indigo",
  ];
  
  // ------------------------------------ Ozon ------------------------------------
  
  // Заполнить список атрибутов и значений Озона
  function fillAttributesAndValuesOzon() {
    // Настройки
    let sheetName = "ТЗ OZON"; // Название рабочего листа
    let attrCln = "A"; // Колонка названий атрибутов
    let attrValCln = "C"; // Колонка значений атрибутов
    let attrCheckCln = "B"; // Колонка пометки значения атрибута
    let startRow = 4; // Номер стартовой строки
    let categoryCell = "C1"; // Ячейка с названием категории
  
    let ignoreAttrs = [
      "Бренд",
      "Серии",
      "Страна-изготовитель",
      "Коммерческий тип",
      "Цвет товара",
      "Название цвета",
      "Rich-контент JSON",
      "Количество заводских упаковок",
      "Озон.Видео: название",
      "Озон.Видео: ссылка ",
      "Озон.Видеообложка: ссылка",
      "Код ТН ВЭД электроника",
    ];
    // ----------
  
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    let dashWB = spreadsheet.getSheetByName("Дашборд WB");
    if (dashWB != null) dashWB.hideSheet();
  
    let ozonWB = spreadsheet.getSheetByName("Дашборд OZON");
    if (ozonWB != null) ozonWB.showSheet().setTabColor("#6495ed");
  
    spreadsheet.getSheetByName("ТЗ OZON").showSheet().setTabColor("#6495ed");
    spreadsheet.getSheetByName("ТЗ WB").hideSheet();
    spreadsheet.getSheetByName("Калькулятор").hideSheet();
    spreadsheet.getSheetByName("Выкупы").hideSheet();
    spreadsheet.getSheetByName("Подкатегории").hideSheet();
  
    let offset = startRow - 1;
    let sh = spreadsheet.getSheetByName(sheetName);
  
    let catName = sh.getRange(categoryCell).getValue();
    let catId = getCategoryId(catName);
  
    if (!catId) throw `Название категории "${catName}" не найдено`;
  
    let attrs = getCategoryAttributes(catId);
    attrs = attrs.filter((v) => !ignoreAttrs.includes(v.name));
  
    let attrsData = [];
    let valuesData = [];
  
    attrs.forEach((atr) => {
      attrsData.push([atr.name]);
      valuesData.push([""]);
  
      let atrValues = getAttrValues(atr.id, catId);
      if (!atrValues) {
        return;
      }
  
      atrValues = atrValues.sort((a, b) => a.value.localeCompare(b.value));
      atrValues.forEach((val, i) => {
        attrsData.push([""]);
        valuesData.push([val.value]);
      });
    });
  
    let diffRows = valuesData.length + offset - sh.getMaxRows();
    if (diffRows > 0) {
      sh.insertRows(sh.getMaxRows(), diffRows);
    }
  
    sh.getRange(`${attrCln}${startRow}:${attrCln}`).clearContent();
    sh.getRange(`${attrValCln}${startRow}:${attrValCln}`).clearContent();
    sh.getRange(`${attrCheckCln}${startRow}:${attrCheckCln}`)
      .clearContent()
      .insertCheckboxes();
  
    let attrsRange = sh.getRange(
      `${attrCln}${startRow}:${attrCln}${offset + attrsData.length}`
    );
    let attrsValsRange = sh.getRange(
      `${attrValCln}${startRow}:${attrValCln}${offset + valuesData.length}`
    );
  
    attrsRange.setValues(attrsData);
    attrsValsRange.setValues(valuesData);
  
    // filter checkboxes
    let dataRange = sh.getRange(`${attrCln}${startRow}:${attrCheckCln}`);
  
    dataRange.getValues().forEach((v, i) => {
      if (v[0])
        sh.getRange(`${attrCheckCln}${i + 1 + offset}`).removeCheckboxes();
    });
  }
  
  function clearOzonFalseAttributes() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ТЗ OZON");
    let range = sheet.getRange("A:B");
    let values = range.getValues();
  
    for (let i = values.length - 1; i >= 3; i--) {
      if (values[i][0] && (i === values.length - 1 || values[i + 1][0])) {
        //sheet.deleteRow(i + 1);
        sheet.hideRows(i + 1);
        values.splice(i, 1);
        continue;
      }
  
      if (!values[i][0] && !values[i][1]) {
        //sheet.deleteRow(i + 1);
        sheet.hideRows(i + 1);
        values.splice(i, 1);
        continue;
      }
    }
  }
  
  // Определить id категории по имени из списка
  function getCategoryId(catName) {
    if (!checkCategoriesSheet()) loadCategories();
  
    let categoriesSheetName = "Categories";
    let list = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(categoriesSheetName)
      .getRange("E:F")
      .getValues();
  
    let find = list.filter((v) => v[0] === catName)[0];
  
    return find ? find[1] : null;
  }
  
  // Проверка наличия списка категорий
  function checkCategoriesSheet() {
    let categoriesSheetName = "Categories";
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(categoriesSheetName);
  
    return sh != null && sh.getLastRow() != 0;
  }
  
  // Выгрузить все категории в лист Categories
  function loadCategories() {
    let categoriesSheetName = "Categories";
  
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(categoriesSheetName);
  
    if (sh == null) {
      sh = ss.insertSheet(categoriesSheetName, ss.getNumSheets());
      sh.hideSheet();
    }
  
    let catData = getCategories();
    let sheetValues = [];
  
    catData.forEach((c) => {
      c.children.forEach((c2) => {
        c2.children.forEach((c3) => {
          sheetValues.push([
            c.title,
            c.category_id,
            c2.title,
            c2.category_id,
            c3.title,
            c3.category_id,
          ]);
        });
      });
    });
  
    let range = sh.getRange("A:F");
    range.clearContent();
  
    let diffRows = sheetValues.length - range.getNumRows();
    if (diffRows > 0) {
      sh.insertRows(1, diffRows);
      range = sh.getRange("A:F");
    }
  
    range.setValues(sheetValues);
  }
  
  // ---------------- Методы запросов ----------------
  
  // Ozon
  
  function getCategories(catId = null) {
    let data =
      catId == null
        ? {}
        : {
          category_id: catId,
          language: "DEFAULT",
        };
    let response = ozonApiRequest(
      "https://api-seller.ozon.ru/v2/category/tree",
      data
    );
    return response.result;
  }
  
  function getCategoryAttributes(catId, attrType = "ALL") {
    let data = {
      attribute_type: attrType,
      category_id: [catId],
      language: "DEFAULT",
    };
    let response = ozonApiRequest(
      "https://api-seller.ozon.ru/v3/category/attribute",
      data
    );
    return response.result[0].attributes;
  }
  
  function getAttrValues(attrId, catId, limit = 500) {
    let data = {
      attribute_id: attrId,
      category_id: catId,
      language: "DEFAULT",
      last_value_id: 0,
      limit: limit,
    };
    return ozonApiRequest(
      "https://api-seller.ozon.ru/v2/category/attribute/values",
      data
    ).result;
  }
  
  function ozonApiRequest(url = "", data = {}) {
    let clientId = "28805";
    let apiKey = "4db8ea99-7b00-43bd-aba1-5bc0b916aade";
  
    let options = {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: {
        "Client-Id": clientId,
        "Api-Key": apiKey,
      },
      payload: JSON.stringify(data),
    };
  
    let response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  }
  
  // Heroku
  
  function parseSkus(query) {
    Logger.log(`Request heroku at wbQuery=${encodeURIComponent(query)}`);
    let response = UrlFetchApp.fetch(
      `https://zmpstat.herokuapp.com/?wbQuery=${encodeURIComponent(query)}`
    );
    Logger.log(`Recieved heroku response`);
    return response.getContentText();
  }
  
  // MpStats
  
  function getSkuListBySearch(query, lastDays = 30) {
    let dateTo = new Date();
    dateTo.setDate(dateTo.getDate() - 1);
    dateTo = dateTo.toISOString().split("T")[0];
  
    let dateFrom = new Date();
    dateFrom.setDate(dateFrom.getDate() - lastDays);
    dateFrom = dateFrom.toISOString().split("T")[0];
  
    let data = {
      dataFilter: {},
      filterModel: {},
      sortModel: [{ colId: "revenue", sort: "desc", sortIndex: 0 }],
      fields: ["id", "category", "revenue"],
      pageModel: { page: 1 },
      tpl: parseSkus(query),
    };
  
    let response = mpStatsRequest(
      `wb/get/search?d1=${dateFrom}&d2=${dateTo}&path=${encodeURI(
        query
      )}&type=json&salesFields=0`,
      data
    );
  
    if (!response.data || !response.data.length) throw "Нет данных";
  
    return response.data.map((v) => v.id);
  }
  
  function skuStringToArray(skuString) {
    return skuString
      .slice(1, -1)
      .split(",")
      .filter((v, i) => i % 2 == 1)
      .map((v) => v.slice(0, -1).trim());
  }
  
  // -------------------------------------------- Калькулятор --------------------------------------------
  
  function calcProcessing() {
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Калькулятор")
      .showSheet();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Выкупы").showSheet();
  
    try {
      const skuCell = "A4"; // адрес ячейки sku
      const nDaysCell = "N1"; // адрес ячейки кол-ва дней
      const nTopCell = "K1"; // адрес ячейки ТОП
  
      const priceCell = "C4"; // адрес ячейки заполнения цены
      const ransomCell = "G4"; // адрес ячейки заполнения выкупа
      const ownSalesCell = "D4"; // адрес ячейки суммы своих продаж
      const avgTopCell = "F4"; // адрес ячейки средней выручки топ
      const logisticCell = "K4"; // адрес ячейки стоимости доставки
      const fboCell = "N4"; // адрес ячейки процента комиссии FBO
  
      const topRange = 5; // границы для вычисления средней выручки
  
      const calcSheetName = "Калькулятор"; // название вкладки
  
      const sh = SpreadsheetApp.getActive().getSheetByName(calcSheetName);
      if (!sh) throw `Не найден лист ${calcSheetName}`;
  
      let isConcurr = false;
  
      let sku = sh.getRange(skuCell).getValue();
      if (!sku || !/^\d+$/.test(sku)) {
        throw "Не введен SKU";
        //let concurrentSku = SpreadsheetApp.getActive()
        //  .getSheetByName("ТЗ WB")
        //  .getRange("G1")
        //  .getValue();
        //if (!/^\d+$/.test(concurrentSku))
        //  throw "Не введен ID предмета при отсутствии SKU";
        //sku = concurrentSku;
        //isConcurr = true;
      }
  
      const nDays = sh.getRange(nDaysCell).getValue();
      if (!nDays) throw `Не указано количество дней`;
  
      const nTop = sh.getRange(nTopCell).getValue();
      if (!nTop) throw `Не указан топ`;
  
      // определение дат (дд.мм.гггг) - по количеству дней
      let dateTo = new Date();
      dateTo.setDate(dateTo.getDate() - 1);
      dateTo = dateTo.toLocaleDateString("ru-RU");
  
      let dateFrom = new Date();
      dateFrom.setDate(dateFrom.getDate() - nDays);
      dateFrom = dateFrom.toLocaleDateString("ru-RU");
  
      // получить данные с мпстатс
      Logger.log("Получение данных цены, выкупа с mpstats через heroku");
      let data = herokuGetMpstatsData(sku, dateFrom, dateTo);
  
      if (!data.price) throw "Ошибка получения цены с mpstats..";
      if (!data.ransom) throw "Ошибка получения выкупа с mpstats..";
      if (!data.fbo) throw "Ошибка получения FBO с mpstats..";
      if (!data.logistic) data.logistic = ""; //throw "Ошибка получения цены доставки с mpstats..";
  
      // заполнить цену, выкуп, доставку, комиссию FBO
      sh.getRange(priceCell).setValue(
        isConcurr ? "ВВЕСТИ" : data.price.replace(/,00$/g, "")
      );
      sh.getRange(ransomCell).setValue(`${data.ransom}%`);
      sh.getRange(logisticCell).setValue(`${data.logistic.replace(/,00$/g, "")}`);
      sh.getRange(fboCell).setValue(`${data.fbo}%`);
  
      // получить данные о выручках по предмету
      Logger.log("Получение данных о выручке с mpstats");
      let revenues = getRevenueList(data.subjId, nDays);
  
      // Посчитать и внести сумму своих продаж
      if (!isConcurr) {
        let sales = getSkuSales(sku, nDays);
        sales = sales
          .map((v) => v.sales * v.final_price)
          .reduce((a, b) => a + b, 0);
        sh.getRange(ownSalesCell).setValue(`=${sales}*G4`);
      } else {
        sh.getRange(ownSalesCell).setValue(0);
      }
  
      // Посчитать и внести среднее по диапазону топ
      let avgTopFrom = nTop > topRange ? nTop - topRange : 1;
      let avgTopTo = nTop < 100 - topRange ? nTop + topRange : 100;
      let revenuesRange = revenues.slice(avgTopFrom - 1, avgTopTo - 1);
      let avgTop = revenuesRange.reduce((a, b) => a + b) / revenuesRange.length;
      avgTop = Math.floor(avgTop);
      sh.getRange(avgTopCell).setValue(avgTop);
      sh.showSheet();
    } catch (err) {
      SpreadsheetApp.getUi().alert(err);
    }
  }
  
  function calcProcessingBySubjId() {
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Калькулятор")
      .showSheet();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Выкупы").showSheet();
    try {
      const subjIdCell = "G1"; // адрес ячейки sku
      const nDaysCell = "N1"; // адрес ячейки кол-ва дней
      const nTopCell = "K1"; // адрес ячейки ТОП
  
      const priceCell = "C4"; // адрес ячейки заполнения цены
      const ransomCell = "G4"; // адрес ячейки заполнения выкупа
      const ownSalesCell = "D4"; // адрес ячейки суммы своих продаж
      const avgTopCell = "F4"; // адрес ячейки средней выручки топ
      const logisticCell = "K4"; // адрес ячейки стоимости доставки
      const fboCell = "N4"; // адрес ячейки процента комиссии FBO
  
      const topRange = 5; // границы для вычисления средней выручки
  
      const calcSheetName = "Калькулятор"; // название вкладки
  
      const sh = SpreadsheetApp.getActive().getSheetByName(calcSheetName);
      if (!sh) throw `Не найден лист ${calcSheetName}`;
  
      let subjectId = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName("ТЗ WB")
        .getRange(subjIdCell)
        .getValue()
        .toString();
  
      if (!subjectId || !/^\d+$/.test(subjectId))
        throw `Не ведён доступный id предмета в ячейке ${subjectId}`;
  
      const nDays = sh.getRange(nDaysCell).getValue();
      if (!nDays) throw `Не указано количество дней`;
  
      const nTop = sh.getRange(nTopCell).getValue();
      if (!nTop) throw `Не указан топ`;
  
      // определение дат (дд.мм.гггг) - по количеству дней
      let dateTo = new Date();
      dateTo.setDate(dateTo.getDate() - 1);
      dateTo = dateTo.toLocaleDateString("ru-RU");
  
      let dateFrom = new Date();
      dateFrom.setDate(dateFrom.getDate() - nDays);
      dateFrom = dateFrom.toLocaleDateString("ru-RU");
  
      let sku = getSkuList(subjectId, null, nDays)[0];
  
      // получить данные с мпстатс
      Logger.log("Получение данных цены, выкупа с mpstats через heroku");
      let data = herokuGetMpstatsData(sku, dateFrom, dateTo);
  
      if (!data.price) throw "Ошибка получения цены с mpstats..";
      if (!data.ransom) throw "Ошибка получения выкупа с mpstats..";
      if (!data.fbo) throw "Ошибка получения FBO с mpstats..";
      if (!data.logistic) data.logistic = ""; //throw "Ошибка получения цены доставки с mpstats..";
  
      // заполнить цену, выкуп, доставку, комиссию FBO
      sh.getRange(priceCell).setValue("ВВЕСТИ");
      sh.getRange(ransomCell).setValue(`${data.ransom}%`);
      sh.getRange(logisticCell).setValue(`${data.logistic.replace(/,00$/g, "")}`);
      sh.getRange(fboCell).setValue(`${data.fbo}%`);
  
      // получить данные о выручках по предмету
      Logger.log("Получение данных о выручке с mpstats");
      let revenues = getRevenueList(data.subjId, nDays);
  
      // Свои продажи = 0
      sh.getRange(ownSalesCell).setValue(0);
  
      // Посчитать и внести среднее по диапазону топ
      let avgTopFrom = nTop > topRange ? nTop - topRange : 1;
      let avgTopTo = nTop < 100 - topRange ? nTop + topRange : 100;
      let revenuesRange = revenues.slice(avgTopFrom - 1, avgTopTo - 1);
      let avgTop = revenuesRange.reduce((a, b) => a + b) / revenuesRange.length;
      avgTop = Math.floor(avgTop);
      sh.getRange(avgTopCell).setValue(avgTop);
      sh.showSheet();
    } catch (err) {
      SpreadsheetApp.getUi().alert(err);
    }
  }
  
 // Заполнить запросы-двойки в выкупах
function fillRansomQueries(nKeys = 5) {
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Выкупы");
  let fillRange = sh.getRange(`E4:F${3 + nKeys}`);
  let fillValues = fillRange.getValues();

  let shQueries = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Запросы");
  let queries = shQueries.getRange(`A2:D${shQueries.getLastRow()}`).getValues();
  queries = queries.filter((v) => v[0] == 2).slice(0, nKeys);

  if (!queries) return;

  queries.forEach((v, i) => {
    fillValues[i][0] = v[1];
    fillValues[i][1] = v[3];
  });

  fillRange.setValues(fillValues);
}
  
  function herokuGetMpstatsData(sku, d1 = null, d2 = null) {
    let url = `https://zmpstat.herokuapp.com/mpstats?sku=${sku}`;
  
    if (d1 && d2) url += `&d1=${d1}&d2=${d2}`;
  
    Logger.log(`Request heroku at ${url}`);
  
    let response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
    }).getContentText();
    Logger.log(`Recieved heroku response`);
  
    if (response === "Can't logging")
      throw "Не удается авторизоваться на mpstats";
  
    if (response === "wrong sku")
      throw "Товар с указанным SKU не найден";
  
    if (response === "Something went wrong") {
      UrlFetchApp.fetch(`https://zmpstat.herokuapp.com/clearmp`);
  
      response = UrlFetchApp.fetch(
        url,
        {
          muteHttpExceptions: true,
        }
      ).getContentText();
  
      if (response === "Something went wrong")
        throw "Ошибка получения данных с mpstats";
    }
  
    let result = JSON.parse(response);
  
    return result;
  }
  
  function getRevenueList(subjectId, lastDays = 5) {
    let dateTo = new Date();
    dateTo.setDate(dateTo.getDate() - 1);
    dateTo = dateTo.toISOString().split("T")[0];
  
    let dateFrom = new Date();
    dateFrom.setDate(dateFrom.getDate() - lastDays);
    dateFrom = dateFrom.toISOString().split("T")[0];
  
    let data = {
      dataFilter: {},
      filterModel: constructFilter(),
      sortModel: [{ colId: "revenue", sort: "desc", sortIndex: 0 }],
      fields: ["id", "category", "revenue"],
    };
  
    let response = mpStatsRequest(
      `wb/get/subject?d1=${dateFrom}&d2=${dateTo}&path=${subjectId}&type=json`,
      data
    );
  
    if (!response.data || !response.data.length)
      throw "Не пришли данные по предмету";
  
    return response.data.map((v) => v.revenue);
  }
  
  function getSkuSales(sku, lastDays = 5) {
  
    let dateTo = new Date();
    dateTo.setDate(dateTo.getDate() - 1);
    dateTo = dateTo.toISOString().split("T")[0];
  
    let dateFrom = new Date();
    dateFrom.setDate(dateFrom.getDate() - lastDays);
    dateFrom = dateFrom.toISOString().split("T")[0];
  
    const apiKey = "62e010f690d6f6.72414271c10c7a6b8873f0b809951a51ebf31e69";
  
    let response, responseText, nTries = 0;
  
    if (reqMode && reqMode === "chain") {
      do {
        Logger.log("request h-mpstat get");
        response = UrlFetchApp.fetch(`https://zmpstat.herokuapp.com/mpstats/api/wb/get/item/${sku}/sales?d1=${dateFrom}&d2=${dateTo}`);
  
        responseText = response.getContentText()
  
        if (responseText.includes('error-pages'))
          Logger.log("error");
  
      } while (responseText.includes('error-pages') && n++ < 3)
  
      Logger.log("finish h-mpstat get");
  
    } else {
      response = UrlFetchApp.fetch(`https://mpstats.io/api/wb/get/item/${sku}/sales?d1=${dateFrom}&d2=${dateTo}`,
        {
          muteHttpExceptions: true,
          headers: {
            "Content-Type": "application/json",
            "X-Mpstats-TOKEN": apiKey,
          },
        }
      );
    }
  
    response = response.getContentText();
  
    return JSON.parse(response);
  }
  
  function pingParser() {
    const dt = new Date();
  
    const hours = dt.getUTCHours() + 3;
    const strDate = dt.toDateString();
    debugger
    if (hours >= 8 && hours <= 19 && !strDate.includes("Sat") && !strDate.includes("Sun"))
      UrlFetchApp.fetch("https://zmpstat.herokuapp.com/mpstats");
  }
  