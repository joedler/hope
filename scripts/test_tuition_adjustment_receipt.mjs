import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync(new URL("../gas/build/Finance_Service.js", import.meta.url), "utf8");
const context = vm.createContext({
  console,
  Utilities: {
    formatDate(value) {
      return value instanceof Date ? value.toISOString().slice(0, 10).replace(/-/g, "/") : String(value || "");
    }
  }
});
vm.runInContext(source, context);

const adjustmentDocument = context.getAdjustmentOnlySettlementDocument({
  recordBase: 0,
  planBase: 0,
  planNext: 0,
  pendingPlanBase: 0,
  adjustments: [{
    paymentDocId: "ADJ_2026_09_001",
    paymentPdf: "https://drive.example/adj",
    paymentStatus: "補收單已產"
  }]
});
assert.equal(adjustmentDocument.docId, "ADJ_2026_09_001");
assert.equal(adjustmentDocument.pdfUrl, "https://drive.example/adj");

const item = {
  total: 0,
  primaryDocId: "",
  sourceDocIds: [],
  method: "",
  category: "",
  dateRaw: "",
  pid: ""
};
const originalRow1 = ["2026/09", "宥興", "課程", "後收", "09/23", "", 600, 1200, "", "", ""];
const originalRow2 = ["2026/09", "宥興", "課程", "後收", "09/30", "", 600, 1200, 2400, "R_2026_09_003", ""];
const adjustmentRow = ["2026/09", "宥興", "課程", "後收", "09/02", "", 600, 1200, 1200, "ADJ_2026_09_001", ""];
context.addTuitionSettlementRowToReceiptItem(item, originalRow1);
context.addTuitionSettlementRowToReceiptItem(item, originalRow2);
context.addTuitionSettlementRowToReceiptItem(item, adjustmentRow);
context.finalizeTuitionReceiptItem(item, "Asia/Taipei");

assert.equal(item.total, 3600, "合併收據應累加課程金額，不可只取最後一列總額");
assert.equal(item.docId, "R_2026_09_003", "合併收據應以原 R 單作為主收據編號");
assert.deepEqual(Array.from(item.sourceDocIds), ["R_2026_09_003", "ADJ_2026_09_001"]);

const sheetData = [
  [],
  originalRow1,
  originalRow2.concat(["", "", "", "", "", ""]),
  adjustmentRow.concat(["", "", "", "", "https://drive.example/adj", "補收單已產"])
];
const documents = context.collectTuitionSettlementDocuments(sheetData, "2026/09", "Asia/Taipei");
assert.equal(documents.length, 2, "同一學生的 R 與 ADJ 必須辨識為兩張獨立繳費文件");
assert.equal(documents[0].total, 2400);
assert.equal(documents[1].total, 1200);

assert.match(source, /data\[i\]\[20\]/, "帳務補救讀取必須檢查學費認列月份");
assert.match(source, /getRange\(rowNumber, 21\)/, "成功寫入後必須回標學費認列月份");

console.log("Tuition adjustment and combined receipt regression tests passed.");
