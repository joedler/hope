import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync(new URL("../gas/build/Finance_Service.js", import.meta.url), "utf8");
const context = vm.createContext({ console });
vm.runInContext(source, context);

const reconcile = context.getTuitionPlanReconciliationMonth;
assert.equal(typeof reconcile, "function", "找不到取消課程核對月份函式");

const timeZone = "Asia/Taipei";
assert.equal(reconcile("已核銷", "2026/07", "", "2026/08", timeZone), "2026/08");
assert.equal(reconcile("取消", "2026/07", "", "2026/08", timeZone), "2026/08");
assert.equal(reconcile("取消", "", "", "2026/08", timeZone), "");
assert.equal(reconcile("取消", "", "2026/04", "2026/03", timeZone), "2026/04");
assert.equal(reconcile("取消", "2026/07", "2026/08", "2026/08", timeZone), "2026/08");

const augustPlans = [
  { status: "取消", tuitionMonth: "2026/07", refundMonth: "", hours: 1.5 },
  { status: "已實際上課", tuitionMonth: "2026/07", refundMonth: "", hours: 1.5 }
];
const prepaidHours = augustPlans.reduce((sum, plan) => {
  const month = reconcile(plan.status, plan.tuitionMonth, plan.refundMonth, "2026/08", timeZone);
  return sum + (month === "2026/08" ? plan.hours : 0);
}, 0);
const actualHours = 3;
const septemberPrepaidHours = 4.5;
const unitFee = 1000;
const total = Math.round((septemberPrepaidHours + actualHours - prepaidHours) * unitFee);

assert.equal(prepaidHours, 3, "已預收後取消的 1.5 小時必須保留在前期預繳核對基數");
assert.equal(total, 4500, "均逸案例應由 6,000 元更正為 4,500 元");

console.log("Tuition cancellation regression tests passed.");
