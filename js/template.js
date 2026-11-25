// =============================================================================
// 보고서 양식(템플릿) 채우기 — 회사 양식 엑셀의 {{자리표시}} 칸에 값만 넣고 서식은 그대로 (외부 라이브러리 없음)
//  양식: 자료 폴더의 templates/*.xlsx (Excel 에서 직접 고침). 칸에 {{매출액}} 처럼 적어 둠
//   - 칸에 자리표시 하나만 있으면 값 그대로 (숫자는 숫자로 → 양식의 통화 · % 서식이 적용됨, 비율은 0.4667 같은 소수)
//   - 글과 섞여 있으면 글자로 바꿔 넣음 (숫자는 1,234 꼴)
//  칸의 꾸밈 · 병합 · 열 너비 · 인쇄 설정 · 수식은 건드리지 않음. 열 때 Excel 이 수식을 다시 계산하도록 표시
//  쓸 수 있는 값: KEYS (docs/TEMPLATE_KEYS.md 와 같게)
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const X = PC.Xlsx || require("./xlsx.js");
  const { num, pyRound } = PC;
  const dec = new TextDecoder();

  // 자리표시 이름 → 설명 (화면의 값 목록 · 문서)
  const KEYS = [
    ["프로젝트ID", "프로젝트 ID"], ["과제명", "프로젝트명"], ["과제내용", "프로젝트 설명"], ["고객사", "고객사"], ["PM", "프로젝트 PM"], ["영업대표", "영업대표"],
    ["프로젝트Type", "프로젝트 Type (프로젝트 개요)"], ["결재1", "결재자 1 (기안 · PM)"], ["결재2", "결재자 2 (검토)"], ["결재3", "결재자 3 (사업팀장)"], ["결재4", "결재자 4 (영업팀장)"], ["결재5", "결재자 5 (본부장)"],
    ["매출액", "계약 매출 (원)"], ["시작일", "시작일 YYYY.MM.DD"], ["종료일", "종료일 YYYY.MM.DD"], ["개월수", "기간 개월 수"], ["귀속부서", "매출 귀속부서 (첫 부서)"],
    ["노무비", "내부 인건비 (OT비용 포함)"], ["OT비용", "OT비용 (프로젝트 전체)"], ["재료비", "재료비 합계"], ["재료비율", "재료비 ÷ 총 원가"], ["외주비", "외주 인건비"], ["경비", "일반 · 인비례 · 기타경비 (AI예비비 제외)"], ["AI예비비", "경비 중 이름이 AI예비비인 항목"], ["총원가", "총 원가"],
    ["노무비율", "노무비 ÷ 총 원가 (소수)"], ["외주비율", "외주비 ÷ 총 원가"], ["경비율", "경비 ÷ 총 원가"], ["AI예비비율", "AI예비비 ÷ 총 원가"], ["총원가율", "1 (100%)"],
    ["매출마진", "매출 − 총 원가"], ["수주마진율", "매출마진 ÷ 매출 (소수)"], ["허들", "목표 허들 이익률 (소수)"], ["Gap", "수주마진율 − 허들 (소수)"],
    ["최근버전이름", "가장 최근 저장한 버전 이름"], ["최근버전매출", "그 버전의 계약 매출"], ["최근버전마진율", "그 버전의 수주마진율 (소수)"],
    ["견적합계", "견적 합계 (인건비 + 기타, 할인 전)"], ["견적할인율", "견적 할인율 (소수)"], ["견적할인액", "할인 금액"], ["견적절사액", "절사 금액"], ["견적공급가", "견적 공급가액 (할인 · 절사 후, VAT 별도)"],
    ["작성일", "출력한 날 YYYY.MM.DD"], ["최종작성일", "프로젝트를 마지막으로 저장한 날 YYYY.MM.DD"],
    // 원가상세 (달은 첫 달부터 M1 … M12 — 1년까지, 넘는 달은 빠짐)
    ["월1 … 월12", "달 머리글 (예: 26.01) — 기간 밖은 빈칸"],
    ["내부인력.이름", "반복 줄: 성명 · 역할 · 영역 · 시스템 · 부서 · 직급 · 제안직무 · 제안등급 · 단가 · 총MM · 금액 · M1 … M12(MM)"],
    ["외주인력.이름", "반복 줄: 성명 · 역할 · 영역 · 시스템 · 부서 · 업체 · 기술등급 · 제안직무 · 제안등급 · 단가 · 총MM · 금액 · M1 … M12(MM)"],
    ["경비항목.이름", "반복 줄: 구분 · 항목 · 방식 · 기준금액 · 금액 · M1 … M12(금액)"],
    ["내부합계.이름 · 외주합계.이름 · 인건비합계.이름", "인원수 · 총MM · 금액 · M1 … M12(MM) · C1 … C12(인건비) — 내부 · 전체 금액은 OT비용 포함"],
    ["재료비항목.이름", "반복 줄: 구분 · 내용 · 금액"],
    ["재료비합계.이름", "금액 · M1 … M12"],
    ["경비합계.이름 · 월원가.이름 · 누적원가.이름", "금액 · M1 … M12 (원가: 내부 · 외주 인건비 + 경비)"]
  ];
  const MAX_MONTHS = 12;   // 양식은 1년까지 (월1 … 월12)
  const MODE = { equal_total: "총액 균등", fill_all: "매달 같은 금액", manual: "월별 직접" };
  const AI_RESERVE = "AI예비비";

  const dot = s => String(s || "").slice(0, 10).replace(/-/g, ".");
  const lastDay = ym => { const [y, m] = ym.split("-").map(Number); return `${ym}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`; };
  const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  // 프로젝트 · 계산 결과 → 자리표시 값. o: { lastSaved: Date (마지막 저장), today: Date }
  function values(project, calc, o = {}) {
    const s = calc.summary, b = s.breakdown, d = calc.details;
    const ai = pyRound(d.general_expenses.concat(d.custom_costs).filter(e => String(e.name || "").replace(/\s/g, "") === AI_RESERVE).reduce((a, e) => a + num(e.total_cost), 0), 2);
    const expense = pyRound(b.general_expense + b.internal_overhead + b.external_overhead + b.custom_cost - ai, 2);
    const total = s.total_cost, ratio = v => (total > 0 ? v / total : 0);
    const months = calc.months, start = String(project.start_date || ""), end = String(project.end_date || "");
    const margin = pyRound(s.revenue - total, 2), mRate = s.revenue > 0 ? margin / s.revenue : 0, hurdle = num(project.target_hurdle_rate) / 100;
    // 가장 최근 저장한 버전 (저장 일시 → 번호 순)
    const vers = (project.versions || []).slice().sort((x, y) => String(x.saved_at || "").localeCompare(String(y.saved_at || "")) || (x.no || 0) - (y.no || 0));
    const last = vers[vers.length - 1];
    let lastRev = "", lastRate = "";
    if (last && last.data) {
      const vc = PC.Calc.calculateProjectCost(PC.Model.normalize(PC.clone(last.data), project.settings), project.settings).summary;
      lastRev = vc.revenue;
      lastRate = vc.revenue > 0 ? (vc.revenue - vc.total_cost) / vc.revenue : 0;
    }
    const ap = (project.settings || {}).approvers || [];
    const today = o.today || new Date(), saved = o.lastSaved || today;
    return {
      프로젝트ID: project.project_id || "", 과제명: project.project_name || "", 과제내용: project.description || "", 고객사: project.client_name || "", PM: project.pm_name || "", 영업대표: project.sales_rep || "",
      프로젝트Type: project.project_type || "", 결재1: ap[0] || "", 결재2: ap[1] || "", 결재3: ap[2] || "", 결재4: ap[3] || "", 결재5: ap[4] || "",
      매출액: s.revenue, 시작일: dot(start.length === 7 ? `${start}-01` : start), 종료일: dot(end.length === 7 ? lastDay(end) : end), 개월수: months.length,
      귀속부서: ((project.settings || {}).departments || []).filter(x => String(x).trim())[0] || "",
      노무비: b.internal_labor, OT비용: b.ot_cost || "", 재료비: b.material || "", 재료비율: b.material ? ratio(b.material) : "", 외주비: b.external_labor, 경비: expense, AI예비비: ai || "", 총원가: total,
      노무비율: ratio(b.internal_labor), 외주비율: ratio(b.external_labor), 경비율: ratio(expense), AI예비비율: ai ? ratio(ai) : "", 총원가율: total > 0 ? 1 : 0,
      매출마진: margin, 수주마진율: mRate, 허들: hurdle, Gap: mRate - hurdle,
      최근버전이름: last ? last.name || "" : "", 최근버전매출: lastRev, 최근버전마진율: lastRate,
      견적합계: calc.estimate.subtotal, 견적할인율: calc.estimate.discount_rate / 100, 견적할인액: calc.estimate.discount, 견적절사액: calc.estimate.truncation, 견적공급가: calc.estimate.supply_total,
      작성일: dot(ymd(today)), 최종작성일: dot(ymd(saved)),
      ...detailValues(project, calc)
    };
  }

  // 원가상세: 사람 · 경비 줄(반복)과 달별 합계 — 첫 달부터 M1 … M12
  function detailValues(project, calc) {
    const d = calc.details, months = calc.months.slice(0, MAX_MONTHS), out = {};
    for (let i = 1; i <= MAX_MONTHS; i++) out[`월${i}`] = months[i - 1] ? `${months[i - 1].slice(2, 4)}.${months[i - 1].slice(5)}` : "";
    const perMonth = (f, prefix) => Object.fromEntries(Array.from({ length: MAX_MONTHS }, (_, i) => [`${prefix || "M"}${i + 1}`, months[i] ? f(months[i]) : ""]));
    const person = (r, kind) => Object.assign({
      성명: r.name || "", 역할: r.role || "", 영역: r.area || "", 시스템: r.system || "", 부서: r.dept || "", 제안직무: r.prop_role || "", 제안등급: r.prop_level || "",
      단가: months.length ? r.calc[months[0]].rate : "", 총MM: r.total_mm, 금액: r.total_cost, 근무지: r.site || ""
    }, kind === "internal" ? { 직급: r.grade || "", 기술등급: r.tech_grade || "" } : { 업체: r.vendor || "", 기술등급: r.grade || "" }, perMonth(m => r.calc[m].mm || ""));
    out.내부인력 = d.internal_labor.map(r => person(r, "internal"));
    out.외주인력 = d.external_labor.map(r => person(r, "external"));
    const sum = (rows, f) => pyRound(rows.reduce((a, r) => a + num(f(r)), 0), 2);
    const laborTotal = rows => Object.assign({ 총MM: sum(rows, r => r.total_mm), 금액: sum(rows, r => r.total_cost) }, perMonth(m => sum(rows, r => r.calc[m].mm)), perMonth(m => sum(rows, r => r.calc[m].cost), "C"));
    const people = rows => rows.filter(r => num(r.total_mm)).length;
    const ot = d.ot_monthly || {};
    out.내부합계 = Object.assign(laborTotal(d.internal_labor), { 인원수: people(d.internal_labor), 금액: calc.summary.breakdown.internal_labor }, perMonth(m => pyRound(sum(d.internal_labor, r => r.calc[m].cost) + num(ot[m]), 2), "C"));   // 금액 · 월 인건비는 OT비용 포함
    out.외주합계 = Object.assign(laborTotal(d.external_labor), { 인원수: people(d.external_labor) });
    const all = d.internal_labor.concat(d.external_labor);
    out.인건비합계 = Object.assign({ 인원수: people(all), 총MM: sum(all, r => r.total_mm), 금액: pyRound(calc.summary.breakdown.internal_labor + calc.summary.breakdown.external_labor, 2) }, perMonth(m => sum(all, r => r.calc[m].mm)), perMonth(m => pyRound(sum(all, r => r.calc[m].cost) + num(ot[m]), 2), "C"));
    out.재료비항목 = (d.materials || []).map(x => ({ 구분: x.category || "", 내용: x.content || "", 금액: x.total_cost }));
    out.재료비합계 = Object.assign({ 금액: calc.summary.breakdown.material || 0 }, perMonth(m => num((d.material_monthly || {})[m])));
    const exp = [
      ...d.general_expenses.map(e => ({ 구분: "일반경비", 항목: e.name || "", 방식: MODE[e.dist_mode] || "", 기준금액: e.dist_mode === "fill_all" ? num(e.unit_input) : e.dist_mode === "manual" ? "" : num(e.total_input), 금액: e.total_cost, m: e.calc_monthly })),
      ...(d.mm_overheads || []).map(o => ({ 구분: o.category === "mm_internal" ? "인비례(내부)" : "인비례(외주)", 항목: o.account, 방식: "MM × 기준금액", 기준금액: o.unit_amount, 금액: o.total_cost, m: o.calc_monthly })),
      ...d.custom_costs.map(e => ({ 구분: "기타경비", 항목: e.name || "", 방식: MODE[e.dist_mode] || "", 기준금액: e.dist_mode === "fill_all" ? num(e.unit_input) : e.dist_mode === "manual" ? "" : num(e.total_input), 금액: e.total_cost, m: e.calc_monthly }))
    ];
    out.경비항목 = exp.map(e => Object.assign({ 구분: e.구분, 항목: e.항목, 방식: e.방식, 기준금액: e.기준금액, 금액: e.금액 }, perMonth(m => num(e.m[m]))));
    const expM = m => sum(exp, e => e.m[m]);
    out.경비합계 = Object.assign({ 금액: sum(exp, e => e.금액) }, perMonth(expM));
    const monthCost = m => pyRound(sum(d.internal_labor, r => r.calc[m].cost) + num(ot[m]) + sum(d.external_labor, r => r.calc[m].cost) + num((d.material_monthly || {})[m]) + expM(m), 2);
    out.월원가 = Object.assign({ 금액: calc.summary.total_cost }, perMonth(monthCost));
    let cum = 0;
    out.누적원가 = Object.assign({ 금액: "" }, perMonth(m => (cum = pyRound(cum + monthCost(m), 2))));
    out.기간초과 = calc.months.length > MAX_MONTHS ? `※ 기간 ${calc.months.length}개월 중 앞 ${MAX_MONTHS}개월만 표시` : "";
    return out;
  }

  const PH = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const asText = v => (typeof v === "number" ? v.toLocaleString("ko-KR", { maximumFractionDigits: 2 }) : String(v === undefined || v === null ? "" : v));
  const cellText = (cXml, type, shared) => {
    if (type === "s") { const m = /<v>(\d+)<\/v>/.exec(cXml); return m ? shared[Number(m[1])] : null; }
    if (type === "inlineStr") return X.xmlUnesc((cXml.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join(""));
    if (type === "str") { const m = /<v>([\s\S]*?)<\/v>/.exec(cXml); return m ? X.xmlUnesc(m[1]) : null; }
    return null;
  };
  const CELL = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
  const ROW = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g;
  const colNo = ref => [...ref].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  // 열 숨기기: <cols> 의 범위를 열 하나씩 나눠 hidden 을 붙임 (없는 열은 새로)
  function hideCols(xml, set) {
    const m = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
    const entries = m ? m[1].match(/<col\b[^>]*\/>/g) || [] : [];
    const out = [], covered = new Set();
    const attrOf = (t, a) => Number((new RegExp(`\\b${a}="(\\d+)"`).exec(t) || [])[1]);
    entries.forEach(t => {
      const a = attrOf(t, "min"), b = attrOf(t, "max");
      let overlap = false;
      for (let c = a; c <= b; c++) { covered.add(c); if (set.has(c)) overlap = true; }
      if (!overlap) return out.push([a, t]);
      for (let c = a; c <= b; c++) {
        let one = t.replace(/\bmin="\d+"/, `min="${c}"`).replace(/\bmax="\d+"/, `max="${c}"`);
        if (set.has(c)) one = one.replace(/\s+hidden="[^"]*"/, "").replace(/\/>$/, ' hidden="1"/>');
        out.push([c, one]);
      }
    });
    set.forEach(c => { if (!covered.has(c)) out.push([c, `<col min="${c}" max="${c}" width="9" customWidth="1" hidden="1"/>`]); });
    const cols = `<cols>${out.sort((x, y) => x[0] - y[0]).map(x => x[1]).join("")}</cols>`;
    return m ? xml.replace(m[0], cols) : xml.replace("<sheetData", `${cols}<sheetData`);
  }
  const rowNo = rowXml => Number((/^<row\b[^>]*\br="(\d+)"/.exec(rowXml) || [])[1] || 0);
  // 줄 번호 바꾸기: 줄 · 칸의 r 속성 (A5 → A7)
  const renumber = (rowXml, n) => rowXml.replace(/^<row\b([^>]*?)\br="\d+"/, `<row$1r="${n}"`).replace(/(<c\b[^>]*?\br=")([A-Z]+)\d+"/g, `$1$2${n}"`);

  // 양식(xlsx)을 채워 새 파일(Blob)로. → { blob, used, unknown(모르는 이름), empty(값이 없는 이름) }
  //  {{이름}}: 값 하나 · {{묶음.이름}}: vals[묶음] 이 객체면 그 안의 값
  //  {{목록.이름}}: vals[목록] 이 배열이면 그 줄을 항목 수만큼 복사해서 채움 (반복 줄 — 아래 줄 · 병합 · 인쇄 영역은 밀림)
  async function fill(src, vals) {
    const buf = src instanceof ArrayBuffer ? src : ArrayBuffer.isView(src) ? src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) : await src.arrayBuffer();
    const files = await X.unzip(buf);
    const names = Object.keys(files);
    const text = n => dec.decode(files[n]);
    const find = n => names.find(x => x.toLowerCase() === n.toLowerCase());
    const ssName = find("xl/sharedStrings.xml");
    const shared = ssName ? (text(ssName).match(/<si\b[\s\S]*?<\/si>/g) || []).map(si => X.xmlUnesc((si.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join(""))) : [];
    const used = new Set(), unknown = new Set(), empty = new Set();
    const has = (o, k) => o && Object.prototype.hasOwnProperty.call(o, k);
    // 자리표시 이름 → 값 (item: 반복 줄의 항목)
    const resolve = (key, list, item) => {
      used.add(key);
      const dotAt = key.indexOf(".");
      let v, known = true;
      if (dotAt > 0) {
        const a = key.slice(0, dotAt), b = key.slice(dotAt + 1);
        if (list && a === list) { known = !item || has(item, b) || has((vals[list] || [])[0], b) || !(vals[list] || []).length; v = item ? item[b] : ""; }
        else if (has(vals, a) && vals[a] && typeof vals[a] === "object" && !Array.isArray(vals[a])) { known = has(vals[a], b); v = vals[a][b]; }
        else known = false;
      } else if (has(vals, key)) v = vals[key];
      else known = false;
      if (!known) { unknown.add(key); return ""; }
      // 값이 없는 칸 알림 — 기간 밖의 달(월N · …MN · …CN)과 반복 줄 · 안내 문구는 빼고
      if ((v === "" || v === null || v === undefined) && !(list && key.startsWith(list + ".")) && !/^월\d+$|\.[MC]\d+$|^기간초과$/.test(key)) empty.add(key);
      return v === null || v === undefined ? "" : v;
    };
    // 칸 하나 채우기
    const fillCell = (cXml, list, item) => {
      const tag = /^<c\b[^>]*?\/?>/.exec(cXml)[0];
      if (/<f\b/.test(cXml)) return cXml;   // 수식 칸은 그대로
      const type = (/\st="([^"]*)"/.exec(tag) || [])[1];
      const t = cellText(cXml, type, shared);
      if (t === null || t === undefined || !t.includes("{{")) return cXml;
      const head = tag.replace(/\s+t="[^"]*"/, "").replace(/\/?>$/, "");   // r · s 등 나머지 속성은 그대로
      const only = /^\s*\{\{\s*([^{}]+?)\s*\}\}\s*$/.exec(t);
      const v = only ? resolve(only[1], list, item) : t.replace(PH, (m, k) => asText(resolve(k, list, item)));
      if (v === "" || v === null || v === undefined) return `${head}/>`;
      if (typeof v === "number" && Number.isFinite(v)) return `${head}><v>${v}</v></c>`;
      return `${head} t="inlineStr"><is><t xml:space="preserve">${X.xmlEsc(v)}</t></is></c>`;
    };
    // 이 줄이 반복 줄인가: {{목록.이름}} 의 목록이 배열인 자리표시가 있으면 그 목록 이름
    const repeatList = rowXml => {
      for (const cXml of rowXml.match(CELL) || []) {
        const tag = /^<c\b[^>]*?\/?>/.exec(cXml)[0], t = cellText(cXml, (/\st="([^"]*)"/.exec(tag) || [])[1], shared);
        if (!t) continue;
        for (const m of t.matchAll(PH)) { const a = m[1].split(".")[0]; if (m[1].includes(".") && Array.isArray(vals[a])) return a; }
      }
      return null;
    };

    // 시트 파일 → 시트 이름 (인쇄 영역 등 이름 정의를 밀 때)
    const wbName = find("xl/workbook.xml"), relName = find("xl/_rels/workbook.xml.rels");
    const rels = {};
    if (relName) (text(relName).match(/<Relationship\b[^>]*>/g) || []).forEach(r => { const id = (/\bId="([^"]+)"/.exec(r) || [])[1], tg = (/\bTarget="([^"]+)"/.exec(r) || [])[1]; if (id && tg) rels[id] = tg.replace(/^\/?xl\//, ""); });
    const sheetOf = {};
    if (wbName) (text(wbName).match(/<sheet\b[^>]*>/g) || []).forEach(t => { const nm = X.xmlUnesc((/\bname="([^"]*)"/.exec(t) || [])[1] || ""), id = (/\br:id="([^"]+)"/.exec(t) || [])[1]; if (id && rels[id]) sheetOf[`xl/${rels[id]}`.toLowerCase()] = nm; });
    const shifts = {};   // 시트 이름 → 줄 번호 바꾸기 함수

    names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/i.test(n)).forEach(n => {
      let xml = text(n);
      if (!xml.includes("{{") && !shared.some(t => t && t.includes("{{"))) return;
      // 기간 밖의 달: 머리글 칸이 {{월N}} 이고 값이 없으면 그 열을 숨김 (열을 지우면 수식 · 병합이 어긋나므로 숨기기로)
      const hide = new Set();
      (xml.match(CELL) || []).forEach(cXml => {
        const tag = /^<c\b[^>]*?\/?>/.exec(cXml)[0], t = cellText(cXml, (/\st="([^"]*)"/.exec(tag) || [])[1], shared);
        const m = t && /^\s*\{\{\s*(월\d+)\s*\}\}\s*$/.exec(t), ref = (/\br="([A-Z]+)\d+"/.exec(tag) || [])[1];
        if (m && ref && has(vals, m[1]) && (vals[m[1]] === "" || vals[m[1]] === null)) hide.add(colNo(ref));
      });
      if (hide.size) xml = hideCols(xml, hide);
      const grow = [];   // [템플릿 줄, 늘어난 줄 수]
      let offset = 0;
      xml = xml.replace(/<sheetData>([\s\S]*?)<\/sheetData>/, (m, body) => `<sheetData>${body.replace(ROW, rowXml => {
        const r = rowNo(rowXml), list = repeatList(rowXml);
        if (!list) { const out = renumber(rowXml, r + offset).replace(CELL, c => fillCell(c)); return out; }
        const items = vals[list].length ? vals[list] : [null];   // 항목이 없으면 빈 줄 하나 (양식의 테두리는 남김)
        const out = items.map((it, j) => renumber(rowXml, r + offset + j).replace(CELL, c => fillCell(c, list, it))).join("");
        grow.push([r, items.length - 1]);
        offset += items.length - 1;
        return out;
      })}</sheetData>`);
      const newRow = old => old + grow.filter(([r]) => r < old).reduce((a, [, k]) => a + k, 0);
      const shiftRef = ref => ref.replace(/(\$?[A-Z]+\$?)(\d+)/g, (m, col, row) => `${col}${newRow(Number(row))}`);
      if (grow.length) {
        // 병합: 반복 줄 안의 병합은 줄마다 복사, 아래 병합은 밀기
        xml = xml.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (m, body) => {
          const refs = (body.match(/<mergeCell\b[^>]*\bref="([^"]+)"/g) || []).map(t => /ref="([^"]+)"/.exec(t)[1]);
          const out = [];
          refs.forEach(ref => {
            const [a, b] = ref.split(":"), ra = Number(/\d+/.exec(a)[0]), rb = Number(/\d+/.exec(b || a)[0]);
            const g = grow.find(([r]) => r === ra && ra === rb);
            if (g) for (let j = 0; j <= g[1]; j++) out.push(`${a.replace(/\d+/, newRow(ra) + j)}:${(b || a).replace(/\d+/, newRow(ra) + j)}`);
            else out.push(shiftRef(ref));
          });
          return `<mergeCells count="${out.length}">${out.map(r => `<mergeCell ref="${r}"/>`).join("")}</mergeCells>`;
        });
        xml = xml.replace(/<dimension\b[^>]*\/>/, "").replace(/\bsqref="([^"]+)"/g, (m, v) => `sqref="${shiftRef(v)}"`);
        const sname = sheetOf[n.toLowerCase()];
        if (sname) shifts[sname] = shiftRef;
      }
      files[n] = new TextEncoder().encode(xml);
    });
    // 통합 문서: 반복 줄로 밀린 시트의 인쇄 영역 등 · 열 때 수식 다시 계산
    if (wbName) {
      let w = text(wbName);
      w = w.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, a, body, c) => {
        const val = X.xmlUnesc(body).replace(/('?)([^'!,]+)\1!([^,]+)/g, (mm, q, sh, ref) => (shifts[sh] ? `${q}${sh}${q}!${shifts[sh](ref)}` : mm));
        return `${a}${X.xmlEsc(val)}${c}`;
      });
      if (/<calcPr\b/.test(w)) w = w.replace(/<calcPr\b([^>]*?)(\/?)>/, (m, a, sl) => `<calcPr${a.replace(/\s+fullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"${sl}>`);
      else if (/<\/definedNames>/.test(w)) w = w.replace("</definedNames>", `</definedNames><calcPr fullCalcOnLoad="1"/>`);
      else w = w.replace("</sheets>", `</sheets><calcPr fullCalcOnLoad="1"/>`);
      files[wbName] = new TextEncoder().encode(w);
    }
    return { blob: X.zip(names.map(n => ({ name: n, data: files[n] }))), used: [...used], unknown: [...unknown], empty: [...empty] };
  }

  // 출력 파일 이름: <양식 이름>_<프로젝트ID>_<프로젝트명>_<최종작성일 YYYYMMDD>.xlsx
  const clean = s => String(s || "").replace(/[\\/:*?"<>|\r\n\t]+/g, "_").replace(/\s+/g, " ").trim();
  const outName = (templateFile, project, lastSaved) =>
    `${clean(String(templateFile).replace(/\.xlsx$/i, ""))}_${clean(project.project_id)}_${clean(project.project_name)}_${ymd(lastSaved || new Date()).replace(/-/g, "")}.xlsx`;
  // 업로드 양식(프로젝트 올리기)은 보고서 양식 목록에서 뺌
  const isReportTemplate = name => /\.xlsx$/i.test(name) && !/^~\$/.test(name) && !/^project_upload_template/i.test(name);

  PC.Template = { KEYS, values, fill, outName, isReportTemplate };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Template;
})(typeof window !== "undefined" ? window : globalThis);
