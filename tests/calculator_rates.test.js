const assert = require("node:assert/strict");
const test = require("node:test");
const { loadRates, parseResponse } = require("../calculator_rates.js");

const headers = [
  "БАНК", "Особые условия", "Программа", "Наличие субсидии", "Срок кредита",
  "Срок действия дисконта", "Размер дисконта к процентной ставке (в процентных пунктах)",
  "Комиссия/ субсидия от суммы кредита, в %.", "Базовая ставка", "Ставка для клиента",
  "Размер ПВ", "Лимит суммы кредита", "ЖК"
];

function row(special) {
  const values = ["Альфа-Банк", special, "Новостройка", "субсидия", "до 10 лет",
    "на весь срок кредитного договора", "5,11%", "6,61%", "19,00%", "13,89%",
    "от 20,1%", "35 000 000", "все"];
  return { c: values.map((value) => ({ v: value })) };
}

function response(rows) {
  return { status: "ok", table: { cols: headers.map((label) => ({ label })), rows } };
}

test("maps formatted values and excludes suspended offers", () => {
  const data = parseResponse(response([row("приостановлено"), row("кредит до 10 млн"), row("кредит более 10 млн")]));
  assert.equal(data.suspended, 1);
  assert.equal(data.rates.length, 2);
  assert.equal(data.rates[0].client_rate, 0.1389);
  assert.equal(data.rates[0].commission, 0.0661);
  assert.equal(data.rates[0].min_pv, 20.1);
  assert.equal(data.rates[0].loan_max, 10000000);
  assert.equal(data.rates[1].loan_min, 10000000);
});

test("uses Google Sheets formatted numeric cells", () => {
  const offer = row("-");
  offer.c[7] = { v: 0.0661, f: "6,61%" };
  offer.c[9] = { v: 0.1389, f: "13,89%" };
  offer.c[11] = { v: 35000000, f: "35 000 000" };
  const rate = parseResponse(response([offer])).rates[0];
  assert.equal(rate.commission, 0.0661);
  assert.equal(rate.client_rate, 0.1389);
  assert.equal(rate.loan_max, 35000000);
});

test("rejects a changed sheet format", () => {
  const changed = response([row("-")]);
  changed.table.cols[7].label = "Другое поле";
  assert.throws(() => parseResponse(changed), /структура листа/);
});

test("loads the public sheet without a service account", async () => {
  const previousDocument = global.document;
  let loadedUrl;
  const head = {
    appendChild(script) {
      loadedUrl = new URL(script.src);
      script.parentNode = this;
      const callback = loadedUrl.searchParams.get("tqx").split("responseHandler:")[1];
      global[callback](response([row("-")]));
    },
    removeChild(script) { script.parentNode = null; }
  };
  global.document = { head, createElement: () => ({}) };
  try {
    const data = await loadRates();
    assert.equal(loadedUrl.hostname, "docs.google.com");
    assert.equal(loadedUrl.searchParams.get("gid"), "1500037557");
    assert.equal(data.rates.length, 1);
  } finally {
    global.document = previousDocument;
  }
});
