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
    ["노무비", "내부 인건비"], ["외주비", "외주 인건비"], ["경비", "일반 · 인비례 · 기타경비 (AI예비비 제외)"], ["AI예비비", "경비 중 이름이 AI예비비인 항목"], ["총원가", "총 원가"],
    ["노무비율", "노무비 ÷ 총 원가 (소수)"], ["외주비율", "외주비 ÷ 총 원가"], ["경비율", "경비 ÷ 총 원가"], ["AI예비비율", "AI예비비 ÷ 총 원가"], ["총원가율", "1 (100%)"],
    ["매출마진", "매출 − 총 원가"], ["수주마진율", "매출마진 ÷ 매출 (소수)"], ["허들", "목표 허들 이익률 (소수)"], ["Gap", "수주마진율 − 허들 (소수)"],
    ["최근버전이름", "가장 최근 저장한 버전 이름"], ["최근버전매출", "그 버전의 계약 매출"], ["최근버전마진율", "그 버전의 수주마진율 (소수)"],
    ["작성일", "출력한 날 YYYY.MM.DD"], ["최종작성일", "프로젝트를 마지막으로 저장한 날 YYYY.MM.DD"]
  ];
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
    const ap = project.approvers || [];
    const today = o.today || new Date(), saved = o.lastSaved || today;
    return {
      프로젝트ID: project.project_id || "", 과제명: project.project_name || "", 과제내용: project.description || "", 고객사: project.client_name || "", PM: project.pm_name || "", 영업대표: project.sales_rep || "",
      프로젝트Type: project.project_type || "", 결재1: ap[0] || "", 결재2: ap[1] || "", 결재3: ap[2] || "", 결재4: ap[3] || "", 결재5: ap[4] || "",
      매출액: s.revenue, 시작일: dot(start.length === 7 ? `${start}-01` : start), 종료일: dot(end.length === 7 ? lastDay(end) : end), 개월수: months.length,
      귀속부서: ((project.settings || {}).departments || []).filter(x => String(x).trim())[0] || "",
      노무비: b.internal_labor, 외주비: b.external_labor, 경비: expense, AI예비비: ai || "", 총원가: total,
      노무비율: ratio(b.internal_labor), 외주비율: ratio(b.external_labor), 경비율: ratio(expense), AI예비비율: ai ? ratio(ai) : "", 총원가율: total > 0 ? 1 : 0,
      매출마진: margin, 수주마진율: mRate, 허들: hurdle, Gap: mRate - hurdle,
      최근버전이름: last ? last.name || "" : "", 최근버전매출: lastRev, 최근버전마진율: lastRate,
      작성일: dot(ymd(today)), 최종작성일: dot(ymd(saved))
    };
  }

  const PH = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const asText = v => (typeof v === "number" ? v.toLocaleString("ko-KR", { maximumFractionDigits: 2 }) : String(v === undefined || v === null ? "" : v));
  const cellText = (cXml, type, shared) => {
    if (type === "s") { const m = /<v>(\d+)<\/v>/.exec(cXml); return m ? shared[Number(m[1])] : null; }
    if (type === "inlineStr") return X.xmlUnesc((cXml.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join(""));
    if (type === "str") { const m = /<v>([\s\S]*?)<\/v>/.exec(cXml); return m ? X.xmlUnesc(m[1]) : null; }
    return null;
  };

  // 양식(xlsx)을 채워 새 파일(Blob)로. → { blob, used: [자리표시], unknown: [모르는 이름], empty: [값이 없는 이름] }
  async function fill(src, vals) {
    const buf = src instanceof ArrayBuffer ? src : ArrayBuffer.isView(src) ? src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) : await src.arrayBuffer();
    const files = await X.unzip(buf);
    const names = Object.keys(files);
    const text = n => dec.decode(files[n]);
    const ssName = names.find(n => n.toLowerCase() === "xl/sharedstrings.xml");
    const shared = ssName ? (text(ssName).match(/<si\b[\s\S]*?<\/si>/g) || []).map(si => X.xmlUnesc((si.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join(""))) : [];
    const used = new Set(), unknown = new Set(), empty = new Set();
    const has = k => Object.prototype.hasOwnProperty.call(vals, k);
    const valueOf = k => { used.add(k); if (!has(k)) { unknown.add(k); return ""; } const v = vals[k]; if (v === "" || v === null || v === undefined) empty.add(k); return v; };
    names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/i.test(n)).forEach(n => {
      const xml = text(n);
      if (!xml.includes("{{") && !shared.some(t => t && t.includes("{{"))) return;
      const out = xml.replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, cXml => {
        const tag = /^<c\b[^>]*?\/?>/.exec(cXml)[0];
        if (/<f\b/.test(cXml)) return cXml;   // 수식 칸은 그대로
        const type = (/\st="([^"]*)"/.exec(tag) || [])[1];
        const t = cellText(cXml, type, shared);
        if (t === null || t === undefined || !t.includes("{{")) return cXml;
        const head = tag.replace(/\s+t="[^"]*"/, "").replace(/\/?>$/, "");   // r · s 등 나머지 속성은 그대로
        const only = /^\s*\{\{\s*([^{}]+?)\s*\}\}\s*$/.exec(t);
        let v;
        if (only) v = valueOf(only[1]);
        else v = t.replace(PH, (m, k) => asText(valueOf(k)));
        if (v === "" || v === null || v === undefined) return `${head}/>`;
        if (typeof v === "number" && Number.isFinite(v)) return `${head}><v>${v}</v></c>`;
        return `${head} t="inlineStr"><is><t xml:space="preserve">${X.xmlEsc(v)}</t></is></c>`;
      });
      files[n] = new TextEncoder().encode(out);
    });
    // 열 때 수식 다시 계산
    const wbName = names.find(n => n.toLowerCase() === "xl/workbook.xml");
    if (wbName) {
      let w = text(wbName);
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
