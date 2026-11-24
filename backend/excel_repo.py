import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Tuple, List
from zipfile import BadZipFile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


class ExcelRepository:
    def __init__(self, base_dir: str = "../data"):
        self.base_path = Path(base_dir)
        self.master_dir = self.base_path / "master"
        self.projects_dir = self.base_path / "projects"
        self.backup_dir = self.base_path / "backups"
        self.deleted_dir = self.backup_dir / "deleted"
        self.reports_dir = self.base_path / "reports"
        self.templates_dir = self.base_path / "templates"
        for p in (
            self.master_dir,
            self.projects_dir,
            self.backup_dir,
            self.deleted_dir,
            self.reports_dir,
            self.templates_dir,
        ):
            p.mkdir(parents=True, exist_ok=True)

    def default_settings(self) -> Dict[str, Any]:
        return {
            "internal_overhead_ratio": 0.15,
            "external_overhead_ratio": 0.05,
            "default_promotion_month": "03",
            "internal_grades": [
                {"grade": "수석", "rate": 8500000},
                {"grade": "책임", "rate": 7200000},
                {"grade": "선임", "rate": 6000000},
                {"grade": "전임", "rate": 4800000},
            ],
            "external_grades": ["특급", "고급", "중급", "초급"],
            "vendor_rates": [
                {"vendor": "파트너A", "grade": "특급", "rate": 9500000},
                {"vendor": "파트너A", "grade": "고급", "rate": 7800000},
                {"vendor": "파트너A", "grade": "중급", "rate": 6200000},
                {"vendor": "파트너B", "grade": "특급", "rate": 9200000},
                {"vendor": "파트너B", "grade": "고급", "rate": 7500000},
                {"vendor": "파트너B", "grade": "중급", "rate": 6000000},
            ],
            "expense_accounts": [
                {"account": "출장비", "category": "general", "desc": "국내/외 출장 경비"},
                {"account": "회의·식대", "category": "general", "desc": "프로젝트 회의 및 업무추진비"},
                {"account": "장비·SW임차", "category": "general", "desc": "개발장비 및 라이선스"},
                {"account": "사전심의·인증비", "category": "custom", "desc": "외부 심의 및 보안인증 수수료"},
                {"account": "예비비", "category": "custom", "desc": "리스크 대응 예비 비용"},
            ],
        }

    def _quarantine_bad_file(self, path: Path, prefix: str) -> None:
        if path.exists():
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            dest = self.backup_dir / f"{prefix}_{ts}.xlsx"
            try:
                shutil.move(str(path), str(dest))
            except Exception:
                try:
                    path.unlink(missing_ok=True)
                except Exception:
                    pass

    def ensure_master_file(self) -> Tuple[Dict[str, Any], Path]:
        master_path = self.master_dir / "master_rates.xlsx"
        if not master_path.exists():
            defaults = self.default_settings()
            self.save_master_settings(defaults, master_path)
            return defaults, master_path

        try:
            wb = openpyxl.load_workbook(master_path, data_only=True)
        except BadZipFile:
            self._quarantine_bad_file(master_path, "bad_master")
            defaults = self.default_settings()
            self.save_master_settings(defaults, master_path)
            return defaults, master_path

        try:
            if "Master_JSON" in wb.sheetnames and wb["Master_JSON"]["A1"].value:
                loaded = json.loads(str(wb["Master_JSON"]["A1"].value))
                merged = {**self.default_settings(), **loaded}
                return merged, master_path

            # 레거시 시트 호환
            settings = self.default_settings()
            if "Overhead_Config" in wb.sheetnames:
                for r in wb["Overhead_Config"].iter_rows(min_row=2, values_only=True):
                    if not r or not r[0]:
                        continue
                    k = str(r[0]).strip()
                    if k == "internal_overhead_ratio" and r[1] is not None:
                        settings["internal_overhead_ratio"] = float(r[1])
                    elif k == "external_overhead_ratio" and r[1] is not None:
                        settings["external_overhead_ratio"] = float(r[1])
                    elif k == "default_promotion_month" and r[1] is not None:
                        settings["default_promotion_month"] = str(r[1]).zfill(2)

            if "External_Rates" in wb.sheetnames:
                vr = []
                for r in wb["External_Rates"].iter_rows(min_row=2, values_only=True):
                    if r and r[0] and r[1] and r[2] is not None:
                        vr.append({"vendor": str(r[0]).strip(), "grade": str(r[1]).strip(), "rate": float(r[2])})
                if vr:
                    settings["vendor_rates"] = vr
            return settings, master_path
        finally:
            wb.close()

    def save_master_settings(self, settings: Dict[str, Any], target_path: Path | None = None) -> Path:
        path = target_path or (self.master_dir / "master_rates.xlsx")
        wb = openpyxl.Workbook()
        try:
            ws_cfg = wb.active
            ws_cfg.title = "Overhead_Config"
            ws_cfg.append(["config_key", "value", "description"])
            ws_cfg.append(["internal_overhead_ratio", float(settings.get("internal_overhead_ratio", 0.15)), "내부 인비례 경비율"])
            ws_cfg.append(["external_overhead_ratio", float(settings.get("external_overhead_ratio", 0.05)), "외주 인비례 경비율"])
            ws_cfg.append(["default_promotion_month", str(settings.get("default_promotion_month", "03")), "내부 기본 승진월(MM)"])

            ws_ig = wb.create_sheet("Internal_Grades")
            ws_ig.append(["grade", "monthly_rate"])
            for g in settings.get("internal_grades") or []:
                ws_ig.append([g.get("grade", ""), float(g.get("rate") or 0)])

            ws_eg = wb.create_sheet("External_Grades")
            ws_eg.append(["grade"])
            for g in settings.get("external_grades") or []:
                ws_eg.append([str(g)])

            ws_vr = wb.create_sheet("External_Rates")
            ws_vr.append(["vendor", "grade", "monthly_rate"])
            for row in settings.get("vendor_rates") or []:
                ws_vr.append([row.get("vendor", ""), row.get("grade", ""), float(row.get("rate") or 0)])

            ws_ea = wb.create_sheet("Expense_Accounts")
            ws_ea.append(["account", "category", "desc"])
            for row in settings.get("expense_accounts") or []:
                ws_ea.append([row.get("account", ""), row.get("category", "general"), row.get("desc", "")])

            ws_json = wb.create_sheet("Master_JSON")
            ws_json["A1"] = json.dumps(settings, ensure_ascii=False)
            ws_json.sheet_state = "hidden"

            wb.save(path)
        finally:
            wb.close()
        return path

    def project_file_path(self, project_id: str) -> Path:
        safe = "".join(c for c in str(project_id or "PRJ-NEW") if c.isalnum() or c in ("-", "_"))
        return self.projects_dir / f"{safe}.xlsx"

    def list_projects(self) -> List[Dict[str, Any]]:
        rows = []
        for f in sorted(self.projects_dir.glob("*.xlsx")):
            try:
                st = f.stat()
                rows.append({
                    "project_id": f.stem,
                    "file_path": str(f.resolve()),
                    "updated_at": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "mtime": st.st_mtime,
                })
            except Exception:
                continue
        return rows

    def delete_project_safe(self, file_path: str) -> str:
        """안전 삭제: 원본 파일을 data/backups/deleted/ 로 이동"""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"삭제할 프로젝트 파일이 없습니다: {file_path}")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = self.deleted_dir / f"{path.stem}_deleted_{ts}.xlsx"
        shutil.move(str(path), str(dest))
        return str(dest.resolve())

    def load_project_and_close(self, file_path: str) -> Tuple[Dict[str, Any], float]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        master, _ = self.ensure_master_file()
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except BadZipFile:
            self._quarantine_bad_file(path, "bad_project")
            data = self.default_project(path.stem, master)
            return data, 0.0

        mtime = path.stat().st_mtime
        try:
            if "State_JSON" in wb.sheetnames and wb["State_JSON"]["A1"].value:
                data = json.loads(str(wb["State_JSON"]["A1"].value))
                if not data.get("settings"):
                    data["settings"] = master
                if "estimate" not in data:
                    data["estimate"] = self.default_project(path.stem, master)["estimate"]
            else:
                data = self.default_project(path.stem, master)
        finally:
            wb.close()
        return data, mtime

    def _styles(self):
        return {
            "hdr_fill": PatternFill("solid", fgColor="334155"),
            "sub_fill": PatternFill("solid", fgColor="F1F5F9"),
            "tot_fill": PatternFill("solid", fgColor="E2E8F0"),
            "ok_fill": PatternFill("solid", fgColor="DCFCE7"),
            "warn_fill": PatternFill("solid", fgColor="FEE2E2"),
            "hdr_font": Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF"),
            "sec_font": Font(name="맑은 고딕", size=11, bold=True, color="0F172A"),
            "body_font": Font(name="맑은 고딕", size=10, color="1E293B"),
            "bold_font": Font(name="맑은 고딕", size=10, bold=True, color="0F172A"),
            "border": Border(
                left=Side(border_style="thin", color="CBD5E1"),
                right=Side(border_style="thin", color="CBD5E1"),
                top=Side(border_style="thin", color="CBD5E1"),
                bottom=Side(border_style="thin", color="CBD5E1"),
            ),
        }

    def save_project_and_close(
        self,
        file_path: str,
        project: Dict[str, Any],
        calc: Dict[str, Any],
        expected_mtime: float | None = None,
        actor: str = "local_user",
    ) -> float:
        path = Path(file_path)
        if path.exists() and expected_mtime is not None and expected_mtime > 0:
            cur_mtime = path.stat().st_mtime
            if abs(cur_mtime - expected_mtime) > 0.01:
                raise RuntimeError("저장 전 파일이 외부에서 먼저 변경되었습니다. 다시 로드 후 진행하세요.")

        if path.exists():
            try:
                openpyxl.load_workbook(path, read_only=True).close()
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(path, self.backup_dir / f"{path.stem}_{ts}.xlsx")
            except BadZipFile:
                self._quarantine_bad_file(path, "bad_before_save")

        st = self._styles()
        months = calc.get("months") or []
        s = calc["summary"]
        b = s["breakdown"]
        settings = project.get("settings") or self.default_settings()

        wb = openpyxl.Workbook()
        try:
            # 1. 01_Project_Info
            ws1 = wb.active
            ws1.title = "01_Project_Info"
            ws1.append(["항목", "내용"])
            for c in ws1[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]

            info_rows = [
                ("프로젝트 ID", project.get("project_id", "")),
                ("프로젝트명", project.get("project_name", "")),
                ("고객사", project.get("client_name", "")),
                ("프로젝트 PM", project.get("pm_name", "")),
                ("영업대표", project.get("sales_rep", "")),
                ("시작일/월", project.get("start_date") or project.get("start_ym", "")),
                ("종료일/월", project.get("end_date") or project.get("end_ym", "")),
                ("프로젝트 설명", project.get("description", "")),
                ("계약 매출(원)", s["revenue"]),
                ("총 원가(원)", s["total_cost"]),
                ("예상 이익(원)", s["profit"]),
                ("예상 이익률(%)", s["margin_rate"]),
                ("목표 허들(%)", s["hurdle_rate"]),
                ("허들 대비(%p)", s["hurdle_gap"]),
                ("내부 인건비(원)", b["internal_labor"]),
                ("외주 인건비(원)", b["external_labor"]),
                ("일반 경비(원)", b["general_expense"]),
                ("내부 인비례경비(원)", b["internal_overhead"]),
                ("외주 인비례경비(원)", b["external_overhead"]),
                ("기타 비용(원)", b["custom_cost"]),
            ]
            for k, v in info_rows:
                ws1.append([k, v])
                r = ws1.max_row
                ws1.cell(r, 1).fill, ws1.cell(r, 1).font, ws1.cell(r, 1).border = st["sub_fill"], st["bold_font"], st["border"]
                c2 = ws1.cell(r, 2)
                c2.font, c2.border = st["body_font"], st["border"]
                if isinstance(v, (int, float)):
                    c2.number_format = "#,##0.00" if "%" in k else "#,##0"
                    c2.alignment = Alignment(horizontal="right")
                elif k == "프로젝트 설명":
                    c2.alignment = Alignment(wrap_text=True, vertical="top")

            # 2. 02_Project_Settings (프로젝트 전용 기준정보)
            ws2 = wb.create_sheet("02_Project_Settings")
            ws2.append(["구분", "항목1(업체/등급/키)", "항목2(등급/유형)", "값(단가/비율/설명)"])
            for c in ws2[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            ws2.append(["인비례설정", "internal_overhead_ratio", "비율", settings.get("internal_overhead_ratio", 0.15)])
            ws2.append(["인비례설정", "external_overhead_ratio", "비율", settings.get("external_overhead_ratio", 0.05)])
            ws2.append(["승진설정", "default_promotion_month", "기본승진월", settings.get("default_promotion_month", "03")])
            for g in settings.get("internal_grades") or []:
                ws2.append(["내부등급", g.get("grade", ""), "기준단가", float(g.get("rate") or 0)])
            for g in settings.get("external_grades") or []:
                ws2.append(["외주등급", str(g), "등급목록", ""])
            for vr in settings.get("vendor_rates") or []:
                ws2.append(["외주업체단가", vr.get("vendor", ""), vr.get("grade", ""), float(vr.get("rate") or 0)])
            for ea in settings.get("expense_accounts") or []:
                ws2.append(["경비계정", ea.get("account", ""), ea.get("category", "general"), ea.get("desc", "")])

            # 3. 03_Internal_Labor
            ws3 = wb.create_sheet("03_Internal_Labor")
            hdr3 = ["성명", "역할", "등급", "기준단가", "인비례포함", "승진월", "승진단가", "총MM"] + months + ["합계금액"]
            ws3.append(hdr3)
            for c in ws3[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for mem in calc["details"]["internal_labor"]:
                p0 = (mem.get("promotions") or [{}])[0]
                row = [
                    mem.get("name", ""), mem.get("role", ""), mem.get("grade", ""),
                    float(mem.get("base_rate") or 0), "Y" if mem.get("overhead_eligible") else "N",
                    p0.get("effective_ym", ""), p0.get("rate", ""), mem.get("total_mm", 0),
                ] + [mem["calc"][m]["mm"] for m in months] + [mem["total_cost"]]
                ws3.append(row)

            # 4. 04_External_Labor
            ws4 = wb.create_sheet("04_External_Labor")
            hdr4 = ["성명/구분", "역할", "외주업체", "등급", "기준단가", "적용단가", "총MM"] + months + ["합계금액"]
            ws4.append(hdr4)
            for c in ws4[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for ext in calc["details"]["external_labor"]:
                row = [
                    ext.get("name", ""), ext.get("role", ""), ext.get("vendor", ""), ext.get("grade", ""),
                    ext.get("default_rate", 0), ext.get("applied_rate", 0), ext.get("total_mm", 0),
                ] + [ext["calc"][m]["mm"] for m in months] + [ext["total_cost"]]
                ws4.append(row)

            # 5. 05_Expenses
            ws5 = wb.create_sheet("05_Expenses")
            hdr5 = ["구분", "경비계정/항목", "분배방식", "입력금액"] + months + ["합계금액"]
            ws5.append(hdr5)
            for c in ws5[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for exp in calc["details"]["general_expenses"]:
                in_val = exp.get("unit_input", 0) if exp.get("dist_mode") == "fill_all" else exp.get("total_input", 0)
                ws5.append(["일반경비", exp.get("name", ""), exp.get("dist_mode", ""), in_val] + [exp["calc_monthly"][m] for m in months] + [exp["total_cost"]])
            ws5.append(["내부인비례(자동)", "내부 인비례성 경비", f"자동({settings.get('internal_overhead_ratio', 0.15)*100:.1f}%)", "-"] + [calc["details"]["internal_overhead_monthly"][m] for m in months] + [b["internal_overhead"]])
            ws5.append(["외주인비례(자동)", "외주 인비례성 경비", f"자동({settings.get('external_overhead_ratio', 0.05)*100:.1f}%)", "-"] + [calc["details"]["external_overhead_monthly"][m] for m in months] + [b["external_overhead"]])
            for item in calc["details"]["custom_costs"]:
                in_val = item.get("unit_input", 0) if item.get("dist_mode") == "fill_all" else item.get("total_input", 0)
                ws5.append(["기타비용", item.get("name", ""), item.get("dist_mode", ""), in_val] + [item["calc_monthly"][m] for m in months] + [item["total_cost"]])

            # 6. 06_Phases
            ws6 = wb.create_sheet("06_Phases")
            hdr6 = ["회차", "구분", "시작월", "종료월", "회차매출", "회차원가", "회차이익률(%)", "누적매출", "누적원가", "잔여매출", "잔여원가"]
            ws6.append(hdr6)
            for c in ws6[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for p in calc["phases"]:
                ws6.append([
                    p["phase_no"], p["title"], p["start_ym"], p["end_ym"],
                    p["revenue"], p["cost"], p["margin_rate"],
                    p["cum_revenue"], p["cum_cost"], p["rem_revenue"], p["rem_cost"],
                ])

            # 7. 07_Estimate
            ws7 = wb.create_sheet("07_Estimate")
            est = calc["estimate"]
            ws7.append(["견적번호", est["estimate_no"], "견적일자", est["estimate_date"], "매출적용여부", "Y" if est["applied_to_revenue"] else "N"])
            ws7.append(["항목명", "규격/등급", "수량", "단가", "금액", "비고"])
            for c in ws7[2]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for it in est["items"]:
                ws7.append([it.get("item_name", ""), it.get("spec", ""), it["qty"], it["unit_price"], it["amount"], it.get("memo", "")])
            ws7.append(["공급가액 합계", "", "", "", est["supply_total"], ""])
            ws7.append(["부가세(10%)", "", "", "", est["vat"], ""])
            ws7.append(["견적 총액(VAT포함)", "", "", "", est["grand_total"], ""])

            # 8. 99_Audit_Log
            ws8 = wb.create_sheet("99_Audit_Log")
            ws8.append(["일시", "작업자", "프로젝트ID", "계약매출", "총원가", "이익률(%)"])
            for c in ws8[1]:
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            ws8.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), actor, project.get("project_id", ""), s["revenue"], s["total_cost"], s["margin_rate"]])

            # 9. Hidden State_JSON for exact reload
            ws_json = wb.create_sheet("State_JSON")
            ws_json["A1"] = json.dumps(project, ensure_ascii=False)
            ws_json.sheet_state = "hidden"

            for ws in (ws1, ws2, ws3, ws4, ws5, ws6, ws7, ws8):
                ws.freeze_panes = "A2"
                for ci in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(ci)].width = 16
            ws1.column_dimensions["A"].width = 22
            ws1.column_dimensions["B"].width = 48

            wb.save(path)
        finally:
            wb.close()

        return path.stat().st_mtime

    def export_report_excel(self, project: Dict[str, Any], calc: Dict[str, Any]) -> str:
        """2시트 보고서 생성: 01_통합보고 + 02_항목별상세"""
        pid = "".join(c for c in str(project.get("project_id") or "PRJ") if c.isalnum() or c in ("-", "_"))
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = self.reports_dir / f"{pid}_원가보고서_{ts}.xlsx"
        st = self._styles()
        s = calc["summary"]
        b = s["breakdown"]
        est = calc["estimate"]
        months = calc["months"]

        wb = openpyxl.Workbook()
        try:
            # ================= Sheet 1: 01_통합보고 =================
            ws1 = wb.active
            ws1.title = "01_통합보고"
            ws1.merge_cells("A1:H1")
            ws1["A1"] = f"착수 전 프로젝트 원가 통합 보고서 - {project.get('project_name', pid)}"
            ws1["A1"].fill = st["hdr_fill"]
            ws1["A1"].font = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
            ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws1.row_dimensions[1].height = 30

            meta = [
                ("프로젝트 ID", project.get("project_id", ""), "고객사", project.get("client_name", "")),
                ("프로젝트 PM", project.get("pm_name", ""), "영업대표", project.get("sales_rep", "")),
                ("수행 기간", f"{project.get('start_date') or project.get('start_ym', '')} ~ {project.get('end_date') or project.get('end_ym', '')}", "견적 공급가액", est["supply_total"]),
                ("프로젝트 설명", project.get("description", ""), "목표 허들(%)", s["hurdle_rate"]),
            ]
            for idx, (k1, v1, k2, v2) in enumerate(meta, start=3):
                ws1.cell(idx, 1, k1).fill = st["sub_fill"]
                ws1.cell(idx, 1).font = st["bold_font"]
                ws1.cell(idx, 2, v1).font = st["body_font"]
                ws1.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
                ws1.cell(idx, 5, k2).fill = st["sub_fill"]
                ws1.cell(idx, 5).font = st["bold_font"]
                c_v2 = ws1.cell(idx, 6, v2)
                c_v2.font = st["body_font"]
                ws1.merge_cells(start_row=idx, start_column=6, end_row=idx, end_column=8)
                if isinstance(v2, (int, float)):
                    c_v2.number_format = "#,##0.00" if "%" in k2 else "#,##0"
                    c_v2.alignment = Alignment(horizontal="right")
                for col in range(1, 9):
                    ws1.cell(idx, col).border = st["border"]
            ws1.row_dimensions[6].height = 36
            ws1["B6"].alignment = Alignment(wrap_text=True, vertical="top")

            # 원가 요약표
            ws1.cell(8, 1, "원가 및 수익성 요약").font = st["sec_font"]
            ws1.cell(9, 1, "구분").fill, ws1.cell(9, 1).font = st["hdr_fill"], st["hdr_font"]
            ws1.cell(9, 2, "금액(원)").fill, ws1.cell(9, 2).font = st["hdr_fill"], st["hdr_font"]
            ws1.cell(9, 3, "매출대비 비중").fill, ws1.cell(9, 3).font = st["hdr_fill"], st["hdr_font"]
            for col in range(1, 4):
                ws1.cell(9, col).border = st["border"]

            ws1.cell(10, 1, "계약 매출").font = st["bold_font"]
            ws1.cell(10, 2, s["revenue"]).number_format = "#,##0"
            ws1.cell(10, 3, "=IF(B10>0, 1, 0)").number_format = "0.0%"
            for col in range(1, 4):
                ws1.cell(10, col).fill = st["sub_fill"]
                ws1.cell(10, col).border = st["border"]

            cost_items = [
                ("내부 인건비", b["internal_labor"]),
                ("외주 인건비", b["external_labor"]),
                ("일반 경비", b["general_expense"]),
                ("내부 인비례 경비", b["internal_overhead"]),
                ("외주 인비례 경비", b["external_overhead"]),
                ("기타 비용", b["custom_cost"]),
            ]
            for offset, (label, amt) in enumerate(cost_items, start=11):
                ws1.cell(offset, 1, label).font = st["body_font"]
                c_a = ws1.cell(offset, 2, amt)
                c_a.font, c_a.number_format = st["body_font"], "#,##0"
                c_r = ws1.cell(offset, 3, f"=IF($B$10>0, B{offset}/$B$10, 0)")
                c_r.font, c_r.number_format = st["body_font"], "0.0%"
                for col in range(1, 4):
                    ws1.cell(offset, col).border = st["border"]

            ws1.cell(17, 1, "총 원가 합계").font = st["bold_font"]
            ws1.cell(17, 2, "=SUM(B11:B16)").number_format = "#,##0"
            ws1.cell(17, 3, "=IF($B$10>0, B17/$B$10, 0)").number_format = "0.0%"
            for col in range(1, 4):
                ws1.cell(17, col).fill = st["tot_fill"]
                ws1.cell(17, col).font = st["bold_font"]
                ws1.cell(17, col).border = st["border"]

            ws1.cell(18, 1, "예상 이익 / 이익률").font = st["bold_font"]
            ws1.cell(18, 2, "=B10-B17").number_format = "#,##0"
            ws1.cell(18, 3, "=IF($B$10>0, B18/$B$10, 0)").number_format = "0.0%"
            res_fill = st["ok_fill"] if s["hurdle_met"] else st["warn_fill"]
            for col in range(1, 4):
                ws1.cell(18, col).fill = res_fill
                ws1.cell(18, col).font = st["bold_font"]
                ws1.cell(18, col).border = st["border"]

            # 차트 삽입
            chart = BarChart()
            chart.type = "col"
            chart.title = "원가 항목별 구성"
            chart.y_axis.title = "금액(원)"
            data_ref = Reference(ws1, min_col=2, min_row=10, max_row=16)
            cats_ref = Reference(ws1, min_col=1, min_row=11, max_row=16)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.legend = None
            chart.width = 16
            chart.height = 9
            ws1.add_chart(chart, "E9")

            # 분할계약 요약표
            ws1.cell(20, 1, "분할계약 현황 (회차 / 누적 / 잔여)").font = st["sec_font"]
            ph_hdr = ["회차", "기간", "회차매출", "회차원가", "이익률(%)", "누적매출", "누적원가", "잔여매출"]
            for ci, h in enumerate(ph_hdr, start=1):
                c = ws1.cell(21, ci, h)
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            for r_off, p in enumerate(calc["phases"], start=22):
                vals = [
                    p["title"], f"{p['start_ym']}~{p['end_ym']}",
                    p["revenue"], p["cost"], p["margin_rate"],
                    p["cum_revenue"], p["cum_cost"], p["rem_revenue"],
                ]
                for ci, val in enumerate(vals, start=1):
                    c = ws1.cell(r_off, ci, val)
                    c.font, c.border = st["body_font"], st["border"]
                    if isinstance(val, (int, float)):
                        c.number_format = "#,##0.00" if ci == 5 else "#,##0"
                        c.alignment = Alignment(horizontal="right")

            for ci in range(1, 9):
                ws1.column_dimensions[get_column_letter(ci)].width = 18

            # ================= Sheet 2: 02_항목별상세 =================
            ws2 = wb.create_sheet("02_항목별상세")
            cur_row = 1

            # Section 1: 내부인력
            ws2.cell(cur_row, 1, "1. 내부인력 투입 및 인건비 상세").font = st["sec_font"]
            cur_row += 1
            h_int = ["성명", "역할", "등급", "기준단가", "인비례", "승진월/단가", "총MM"] + months + ["합계금액"]
            for ci, h in enumerate(h_int, start=1):
                c = ws2.cell(cur_row, ci, h)
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            cur_row += 1
            int_start = cur_row
            for mem in calc["details"]["internal_labor"]:
                p0 = (mem.get("promotions") or [{}])[0]
                p_str = f"{p0.get('effective_ym', '')}/{p0.get('rate', '')}" if p0.get("effective_ym") else "-"
                vals = [
                    mem.get("name", ""), mem.get("role", ""), mem.get("grade", ""),
                    float(mem.get("base_rate") or 0), "Y" if mem.get("overhead_eligible") else "N",
                    p_str, mem.get("total_mm", 0),
                ] + [mem["calc"][m]["cost"] for m in months] + [mem["total_cost"]]
                for ci, val in enumerate(vals, start=1):
                    c = ws2.cell(cur_row, ci, val)
                    c.font, c.border = st["body_font"], st["border"]
                    if isinstance(val, (int, float)):
                        c.number_format = "#,##0.00" if ci == 7 else "#,##0"
                        c.alignment = Alignment(horizontal="right")
                cur_row += 1
            tot_col = len(h_int)
            tot_letter = get_column_letter(tot_col)
            ws2.cell(cur_row, 1, "내부인건비 소계").font = st["bold_font"]
            ws2.cell(cur_row, tot_col, f"=SUM({tot_letter}{int_start}:{tot_letter}{cur_row-1})" if cur_row > int_start else 0).number_format = "#,##0"
            for ci in range(1, tot_col + 1):
                ws2.cell(cur_row, ci).fill, ws2.cell(cur_row, ci).border = st["tot_fill"], st["border"]
            cur_row += 2

            # Section 2: 외부(외주)인력
            ws2.cell(cur_row, 1, "2. 외부(외주)인력 투입 및 인건비 상세").font = st["sec_font"]
            cur_row += 1
            h_ext = ["성명/구분", "역할", "외주업체", "등급", "기준단가", "적용단가", "총MM"] + months + ["합계금액"]
            for ci, h in enumerate(h_ext, start=1):
                c = ws2.cell(cur_row, ci, h)
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            cur_row += 1
            ext_start = cur_row
            for ext in calc["details"]["external_labor"]:
                vals = [
                    ext.get("name", ""), ext.get("role", ""), ext.get("vendor", ""), ext.get("grade", ""),
                    ext.get("default_rate", 0), ext.get("applied_rate", 0), ext.get("total_mm", 0),
                ] + [ext["calc"][m]["cost"] for m in months] + [ext["total_cost"]]
                for ci, val in enumerate(vals, start=1):
                    c = ws2.cell(cur_row, ci, val)
                    c.font, c.border = st["body_font"], st["border"]
                    if isinstance(val, (int, float)):
                        c.number_format = "#,##0.00" if ci == 7 else "#,##0"
                        c.alignment = Alignment(horizontal="right")
                cur_row += 1
            tot_col_e = len(h_ext)
            tot_letter_e = get_column_letter(tot_col_e)
            ws2.cell(cur_row, 1, "외주인건비 소계").font = st["bold_font"]
            ws2.cell(cur_row, tot_col_e, f"=SUM({tot_letter_e}{ext_start}:{tot_letter_e}{cur_row-1})" if cur_row > ext_start else 0).number_format = "#,##0"
            for ci in range(1, tot_col_e + 1):
                ws2.cell(cur_row, ci).fill, ws2.cell(cur_row, ci).border = st["tot_fill"], st["border"]
            cur_row += 2

            # Section 3: 경비 및 기타비용
            ws2.cell(cur_row, 1, "3. 경비 및 기타비용 상세 (인비례 경비 포함)").font = st["sec_font"]
            cur_row += 1
            h_exp = ["구분", "경비계정/항목", "분배방식", "비고", "", "", ""] + months + ["합계금액"]
            for ci, h in enumerate(h_exp, start=1):
                c = ws2.cell(cur_row, ci, h)
                c.fill, c.font, c.border = st["hdr_fill"], st["hdr_font"], st["border"]
            cur_row += 1
            exp_start = cur_row

            exp_rows = []
            for e in calc["details"]["general_expenses"]:
                exp_rows.append(("일반경비", e.get("name", ""), e.get("dist_mode", ""), e.get("memo", ""), [e["calc_monthly"][m] for m in months], e["total_cost"]))
            exp_rows.append(("내부인비례", "내부 인비례성 경비", "자동계산", "대상인력 연동", [calc["details"]["internal_overhead_monthly"][m] for m in months], b["internal_overhead"]))
            exp_rows.append(("외주인비례", "외주 인비례성 경비", "자동계산", "외주인건비 연동", [calc["details"]["external_overhead_monthly"][m] for m in months], b["external_overhead"]))
            for c_item in calc["details"]["custom_costs"]:
                exp_rows.append(("기타비용", c_item.get("name", ""), c_item.get("dist_mode", ""), c_item.get("memo", ""), [c_item["calc_monthly"][m] for m in months], c_item["total_cost"]))

            for cat, nm, dm, mm_note, m_vals, tot_v in exp_rows:
                vals = [cat, nm, dm, mm_note, "", "", ""] + m_vals + [tot_v]
                for ci, val in enumerate(vals, start=1):
                    c = ws2.cell(cur_row, ci, val)
                    c.font, c.border = st["body_font"], st["border"]
                    if isinstance(val, (int, float)):
                        c.number_format = "#,##0"
                        c.alignment = Alignment(horizontal="right")
                cur_row += 1

            tot_col_x = len(h_exp)
            tot_letter_x = get_column_letter(tot_col_x)
            ws2.cell(cur_row, 1, "경비·기타 소계").font = st["bold_font"]
            ws2.cell(cur_row, tot_col_x, f"=SUM({tot_letter_x}{exp_start}:{tot_letter_x}{cur_row-1})").number_format = "#,##0"
            for ci in range(1, tot_col_x + 1):
                ws2.cell(cur_row, ci).fill, ws2.cell(cur_row, ci).border = st["tot_fill"], st["border"]

            for ci in range(1, ws2.max_column + 1):
                ws2.column_dimensions[get_column_letter(ci)].width = 15

            wb.save(out_path)
        finally:
            wb.close()
        return str(out_path.resolve())

    def parse_uploaded_excel(self, content: bytes) -> Dict[str, Any]:
        master, _ = self.ensure_master_file()
        tmp_path = self.backup_dir / f"_upload_{datetime.now().strftime('%H%M%S%f')}.xlsx"
        tmp_path.write_bytes(content)
        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        try:
            if "State_JSON" in wb.sheetnames and wb["State_JSON"]["A1"].value:
                data = json.loads(str(wb["State_JSON"]["A1"].value))
                if not data.get("settings"):
                    data["settings"] = master
                return data

            ws = wb.active
            kv: Dict[str, Any] = {}
            for row in ws.iter_rows(values_only=True):
                if row and row[0] is not None:
                    kv[str(row[0]).strip()] = row[1]

            proj = self.default_project(str(kv.get("project_id") or "PRJ-UPLOAD"), master)
            proj["project_name"] = str(kv.get("project_name") or proj["project_name"])
            proj["client_name"] = str(kv.get("client_name") or "")
            proj["pm_name"] = str(kv.get("pm_name") or "")
            proj["sales_rep"] = str(kv.get("sales_rep") or "")
            proj["description"] = str(kv.get("description") or "")
            proj["start_date"] = str(kv.get("start_date") or kv.get("start_ym") or proj["start_date"])
            proj["end_date"] = str(kv.get("end_date") or kv.get("end_ym") or proj["end_date"])
            proj["contract_revenue"] = float(kv.get("contract_revenue") or 0.0)
            proj["target_hurdle_rate"] = float(kv.get("target_hurdle_rate") or 18.0)
            return proj
        finally:
            wb.close()
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    def default_project(self, project_id: str, settings: Dict[str, Any] | None = None) -> Dict[str, Any]:
        return {
            "project_id": project_id,
            "project_name": f"{project_id} 프로젝트",
            "client_name": "",
            "pm_name": "",
            "sales_rep": "",
            "description": "",
            "date_mode": "month",
            "start_date": "2026-01",
            "end_date": "2026-06",
            "contract_revenue": 240000000,
            "target_hurdle_rate": 18.0,
            "settings": settings or self.default_settings(),
            "internal_labor": [],
            "external_labor": [],
            "general_expenses": [],
            "custom_costs": [],
            "phases": [],
            "estimate": {
                "estimate_no": "",
                "estimate_date": "",
                "note": "",
                "applied_to_revenue": False,
                "items": [],
            },
        }
