import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const financeSource = fs.readFileSync(new URL("../gas/build/Finance_Service.js", import.meta.url), "utf8");
const finance = vm.createContext({ console });
vm.runInContext(financeSource, finance);

assert.equal(finance.getStoredAssociationSupportAmount_([,,,,,,,,,,,, "是", -1500], 12, 13, 0), 1500);
assert.equal(finance.getStoredAssociationSupportAmount_([,,,,,,,,,,,, "否", -1500], 12, 13, 0), 0);
assert.equal(finance.calculateAssociationSupportAdjustment_({ mode: "後收", supportRecordTotal: 1500 }), -1500);
assert.equal(finance.calculateAssociationSupportAdjustment_({ mode: "預收", supportRecordTotal: 1000, supportPlanBaseTotal: 1000, supportPlanNextTotal: 0, pendingPlanBase: 0 }), 0);
assert.equal(finance.calculateAssociationSupportAdjustment_({ mode: "預收", supportRecordTotal: 0, supportPlanBaseTotal: 1000, supportPlanNextTotal: 0, pendingPlanBase: 0 }), 1000);
assert.equal(finance.calculateAssociationSupportAdjustment_({ mode: "預收", supportRecordTotal: 0, supportPlanBaseTotal: 1000, supportPlanNextTotal: 1500, pendingPlanBase: 1 }), -1500);

const receiptItem = { total: 0, primaryDocId: "", sourceDocIds: [], method: "", category: "", dateRaw: "", pid: "" };
finance.addTuitionSettlementRowToReceiptItem(receiptItem, ["2026/10", "學生", "個別課程", "後收", "", "", 0, 2400, "", "R_2026_10_001"]);
finance.addTuitionSettlementRowToReceiptItem(receiptItem, ["2026/10", "學生", "團體課程", "後收", "", "", 0, 1800, "", ""]);
finance.addTuitionSettlementRowToReceiptItem(receiptItem, ["2026/10", "學生", "與諮詢師有約", "後收", "", "", 0, 1500, "", ""]);
finance.addTuitionSettlementRowToReceiptItem(receiptItem, ["2026/10", "學生", "撐出空間協會支持", "折抵", "", "來源課程：與諮詢師有約；等額折抵", 0, -1500, 4200, "R_2026_10_001"]);
assert.equal(receiptItem.total, 4200, "收據必須保留其他課程，不可因協會支持把整張收據歸零");
assert.equal(finance.getTuitionSavedRowSelectionKey_(["2026/10", "學生", "撐出空間協會支持", "折抵", "", "來源課程：與諮詢師有約；等額折抵"]), "學生::與諮詢師有約");

const coreSource = fs.readFileSync(new URL("../gas/build/Core_Service.js", import.meta.url), "utf8");
const core = vm.createContext({ console });
vm.runInContext(coreSource, core);
assert.equal(core.isAssociationSupportEnabled_("是"), true);
assert.equal(core.isAssociationSupportEnabled_("否"), false);
assert.equal(core.parseAssociationSupportRequest_("true"), true);
const lessons = core.parseAdminProxyLessons_({
  lessons: JSON.stringify([
    { student: "學生", subject: "課程A", date: "2026-10-01", startTime: "13:00", endTime: "14:00", associationSupport: true },
    { student: "學生", subject: "課程B", date: "2026-10-01", startTime: "14:00", endTime: "15:00", associationSupport: false }
  ])
});
assert.equal(lessons.length, 2, "行政代登應接受同一學生多筆課程");
assert.equal(lessons[0].associationSupport, true);

assert.match(financeSource, /const payRate = conf\.fee \* conf\.ratio;/, "講師鐘點費仍須使用原單價乘分潤比例");

const frontendSource = fs.readFileSync(new URL("../docs/index.html", import.meta.url), "utf8");
assert.match(frontendSource, /lessonError\(lesson, this\.adminCourseProxy\.student, this\.adminCourseProxy\.action\)/, "行政代操作檢核必須使用行政選取的學生與操作類型");
assert.match(frontendSource, /x-show="adminProxyLessonError\(lesson\)"/, "行政代操作每筆課程必須顯示停用原因");

console.log("Association support and admin batch registration regression tests passed.");
