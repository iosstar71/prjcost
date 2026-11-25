// =============================================================================
// 계산 회귀 시험의 기준(tests/golden.json)을 지금 JS 계산으로 다시 만든다 — node tools/make_golden.js
//  계산 규칙을 일부러 바꿨을 때만 돌린다 (바뀐 결과가 맞는지 확인한 뒤). 입력 경우는 golden.json 안에 그대로 둠
// =============================================================================
"use strict";
const fs = require("fs"), path = require("path");
const PC = require("../js/util.js");
require("../js/calc.js");
const file = path.join(__dirname, "..", "tests", "golden.json");
const cases = JSON.parse(fs.readFileSync(file, "utf8")).map(g => {
  const expected = PC.Calc.calculateProjectCost(JSON.parse(JSON.stringify(g.input)), g.fallback);
  delete expected.computed_at;
  return { input: g.input, fallback: g.fallback, expected };
});
fs.writeFileSync(file, JSON.stringify(cases, null, 1));
console.log(`[OK] ${cases.length} cases -> tests/golden.json`);
