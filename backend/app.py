from pathlib import Path
from typing import Dict, Any, Optional
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from calculator import calculate_project_cost
from excel_repo import ExcelRepository

APP_VERSION = "v4.0"

app = FastAPI(title=f"Project Cost Local App {APP_VERSION}")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

repo = ExcelRepository(str((Path(__file__).resolve().parent.parent / "data")))


class FilePathReq(BaseModel):
    file_path: str


class ProjectReq(BaseModel):
    file_path: Optional[str] = None
    project: Dict[str, Any]
    expected_mtime: Optional[float] = None


class MasterSettingsReq(BaseModel):
    settings: Dict[str, Any]


@app.get("/")
def index():
    ui_path = Path(__file__).resolve().parent.parent / "frontend" / "index.html"
    return FileResponse(ui_path)


@app.get("/api/bootstrap")
def bootstrap():
    master, _ = repo.ensure_master_file()
    return {
        "version": APP_VERSION,
        "projects": repo.list_projects(),
        "default_settings": master,
    }


@app.post("/api/master-settings")
def save_as_default_settings(req: MasterSettingsReq):
    try:
        repo.save_master_settings(req.settings)
        return {"ok": True, "message": "신규 프로젝트 기본 기준정보로 저장되었습니다."}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/load")
def load_project(req: FilePathReq):
    try:
        master, _ = repo.ensure_master_file()
        project, mtime = repo.load_project_and_close(req.file_path)
        calc = calculate_project_cost(project, master)
        return {"project": project, "calc": calc, "mtime": mtime}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/delete")
def delete_project(req: FilePathReq):
    try:
        moved_path = repo.delete_project_safe(req.file_path)
        return {"deleted": True, "backup_path": moved_path, "projects": repo.list_projects()}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/recalc")
def recalc(req: ProjectReq):
    master, _ = repo.ensure_master_file()
    calc = calculate_project_cost(req.project, master)
    return {"calc": calc}


@app.post("/api/save")
def save_project(req: ProjectReq):
    try:
        master, _ = repo.ensure_master_file()
        calc = calculate_project_cost(req.project, master)
        pid = req.project.get("project_id") or "PRJ-NEW"
        target_path = req.file_path or str(repo.project_file_path(pid).resolve())
        new_mtime = repo.save_project_and_close(
            file_path=target_path,
            project=req.project,
            calc=calc,
            expected_mtime=req.expected_mtime,
        )
        return {
            "saved_path": target_path,
            "mtime": new_mtime,
            "calc": calc,
            "projects": repo.list_projects(),
        }
    except Exception as e:
        raise HTTPException(status_code=409, detail=str(e))


@app.post("/api/export-report")
def export_report(req: ProjectReq):
    master, _ = repo.ensure_master_file()
    calc = calculate_project_cost(req.project, master)
    out_path = repo.export_report_excel(req.project, calc)
    return {"report_path": out_path}


@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        master, _ = repo.ensure_master_file()
        content = await file.read()
        project = repo.parse_uploaded_excel(content)
        calc = calculate_project_cost(project, master)
        return {"project": project, "calc": calc}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
