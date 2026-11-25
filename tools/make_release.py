# -*- coding: utf-8 -*-
"""배포 zip 만들기 — python tools/make_release.py <출력 폴더>

  zip 안: PM-Hub-Cost_v<버전>/ index.html · assets · js · data(master · templates, 빈 projects · backups) · README.md · RELEASE_NOTES.md
  시험 · 도구 · 파이썬 원본은 넣지 않음
"""
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main():
    out_dir = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "dist")
    out_dir.mkdir(parents=True, exist_ok=True)
    version = re.search(r'VERSION = "v([\d.]+)"', (ROOT / "js" / "app.js").read_text(encoding="utf-8")).group(1)
    top = f"PM-Hub-Cost_v{version}"
    zpath = out_dir / f"{top}.zip"
    files = [ROOT / "index.html", ROOT / "README.md", ROOT / "RELEASE_NOTES.md"]
    files += sorted((ROOT / "assets").rglob("*")) + sorted((ROOT / "js").rglob("*.js"))
    files += [ROOT / "data" / "master" / "master_rates.xlsx", ROOT / "data" / "templates" / "project_upload_template.xlsx", ROOT / "data" / "templates" / "Offline사업검토서.xlsx"]
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if f.is_file():
                z.write(f, f"{top}/{f.relative_to(ROOT).as_posix()}")
        for empty in ("data/projects/", "data/backups/deleted/"):
            z.writestr(f"{top}/{empty}", "")
    print(f"[OK] {zpath} ({zpath.stat().st_size:,} bytes, {len(files)} files)")


if __name__ == "__main__":
    main()
