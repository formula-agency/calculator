(function (root) {
  "use strict";

  var SPREADSHEET_ID = "1kooRYR745pcOqj1SbIx3s1McN8I4CaoE5zAcLRuIibQ";
  var SHEET_ID = "1500037557";
  var SHEET_TITLE = "Актуальный (расширенный)";
  var EXPECTED_HEADERS = [
    "БАНК", "Особые условия", "Программа", "Наличие субсидии", "Срок кредита",
    "Срок действия дисконта", "Размер дисконта к процентной ставке", "Комиссия/ субсидия",
    "Базовая ставка", "Ставка для клиента", "Размер ПВ", "Лимит суммы кредита", "ЖК"
  ];
  var requestId = 0;

  function cleanText(value) {
    return String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  }

  function number(value) {
    var match = cleanText(value).replace(/\s/g, "").replace(",", ".").match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) : NaN;
  }

  function rate(value) {
    var parsed = number(value);
    return Number.isFinite(parsed) ? (cleanText(value).includes("%") || parsed > 1 ? parsed / 100 : parsed) : NaN;
  }

  function creditLimits(special, maxInSheet) {
    var lower = special.toLowerCase().replace(",", ".");
    var match = lower.match(/(\d+(?:\.\d+)?)\s*млн/);
    var min = null;
    var max = Number.isFinite(maxInSheet) && maxInSheet > 0 ? maxInSheet : null;
    if (match) {
      var amount = Number(match[1]) * 1000000;
      if (/(?:^|[\s(])до\s/.test(lower)) max = max === null ? amount : Math.min(max, amount);
      else if (/более|свыше|(?:^|[\s(])от\s/.test(lower)) min = amount;
    }
    return { min: min, max: max };
  }

  function parseResponse(response) {
    if (!response || response.status !== "ok" || !response.table || !Array.isArray(response.table.rows)) {
      throw new Error("Google Таблицы не вернули данные ставок.");
    }
    var columns = response.table.cols || [];
    EXPECTED_HEADERS.forEach(function (header, index) {
      if (!columns[index] || !cleanText(columns[index].label).startsWith(header)) {
        throw new Error("Изменилась структура листа ставок (столбец " + (index + 1) + ").");
      }
    });

    var rates = [];
    var suspended = 0;
    response.table.rows.forEach(function (entry) {
      var row = (entry.c || []).map(function (cell) {
        return cell ? (cell.f == null ? cell.v : cell.f) : "";
      });
      var special = cleanText(row[1]);
      if (special.toLowerCase().includes("приостановлено")) {
        suspended++;
        return;
      }
      var bank = cleanText(row[0]);
      var program = cleanText(row[2]);
      var discountTerm = cleanText(row[5]);
      var clientRate = rate(row[9]);
      var commission = rate(row[7]);
      if (!bank && !program && !discountTerm) return;
      if (!bank || !program || !discountTerm || !Number.isFinite(clientRate) ||
          !Number.isFinite(commission) || clientRate < 0 || commission < 0) {
        throw new Error("На актуальном листе есть неполная строка ставок.");
      }
      var minPv = number(row[10]);
      var limits = creditLimits(special, number(row[11]));
      rates.push({
        bank: bank,
        special: special,
        program: program,
        subsidy: cleanText(row[3]),
        loan_term_rule: cleanText(row[4]),
        discount_term: discountTerm,
        discount_delta: rate(row[6]),
        base_rate: rate(row[8]),
        client_rate: clientRate,
        commission: commission,
        pv_text: cleanText(row[10]),
        min_pv: Number.isFinite(minPv) ? minPv : null,
        loan_min: limits.min,
        loan_max: limits.max,
        complex: cleanText(row[12])
      });
    });
    if (!rates.length) throw new Error("На актуальном листе нет действующих ставок.");
    return { rates: rates, suspended: suspended, source: SHEET_TITLE };
  }

  function loadRates() {
    if (typeof document === "undefined") {
      return Promise.reject(new Error("Для загрузки ставок нужен браузер."));
    }
    return new Promise(function (resolve, reject) {
      var callbackName = "__formulaRatesCallback" + Date.now() + "_" + (++requestId);
      var script = document.createElement("script");
      var timeout = setTimeout(function () {
        finish(new Error("Время ожидания Google Таблицы истекло."));
      }, 30000);

      function finish(error, data) {
        clearTimeout(timeout);
        delete root[callbackName];
        if (script.parentNode) script.parentNode.removeChild(script);
        if (error) reject(error);
        else resolve(data);
      }

      root[callbackName] = function (response) {
        try {
          finish(null, parseResponse(response));
        } catch (error) {
          finish(error);
        }
      };
      script.onerror = function () {
        finish(new Error("Не удалось загрузить ставки из Google Таблицы."));
      };
      script.async = true;
      script.src = "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID +
        "/gviz/tq?gid=" + SHEET_ID + "&tqx=" + encodeURIComponent("out:json;responseHandler:" + callbackName);
      document.head.appendChild(script);
    });
  }

  var api = { loadRates: loadRates, parseResponse: parseResponse };
  root.FORMULA_RATES = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
