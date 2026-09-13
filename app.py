"""
TallyDataX - Web Application & Local API Server
FastAPI backend powering live Tally XML extraction, Excel generation, and Excel file merging.
"""

import os
import sys
import re
import shutil
import tempfile
import threading
import webbrowser
import multiprocessing
from typing import List, Dict, Any, Optional

# Ensure freeze support on Windows when packaged with PyInstaller
multiprocessing.freeze_support()

import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import tally_client
import report_engine

app = FastAPI(title="TallyDataX Server", version="1.0.0")

# Enable CORS for local cross-origin if needed
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def get_asset_path(filename: str) -> str:
    """Resolves static assets reliably in development and PyInstaller standalone .exe."""
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        beside = os.path.join(exe_dir, filename)
        if os.path.exists(beside):
            return beside
        meipass_dir = getattr(sys, '_MEIPASS', exe_dir)
        bundled = os.path.join(meipass_dir, filename)
        if os.path.exists(bundled):
            return bundled
    src_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(src_dir, filename)


def get_index_html_path() -> str:
    """Resolves index.html path reliably in development and PyInstaller standalone .exe."""
    return get_asset_path("index.html")


class VoucherQuery(BaseModel):
    from_date: str = "2025-04-01"
    to_date: str = "2026-03-31"
    group_name: str = "Sundry Creditors"
    selected_ledgers: Optional[List[str]] = None


class ExportRequest(BaseModel):
    vouchers: List[Dict[str, Any]]
    party_name: str = "Special Blasts Limited"


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    """Serves the main application UI."""
    html_path = get_index_html_path()
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>index.html not found</h1>", status_code=404)


@app.get("/favicon.ico")
async def serve_favicon():
    """Serves application icon for browser tab."""
    ico_path = get_asset_path("app_icon.ico")
    if os.path.exists(ico_path):
        return FileResponse(ico_path, media_type="image/x-icon")
    return HTMLResponse(status_code=404)


@app.get("/app_icon.png")
async def serve_app_icon_png():
    """Serves high-resolution app icon PNG for web UI."""
    png_path = get_asset_path("app_icon.png")
    if os.path.exists(png_path):
        return FileResponse(png_path, media_type="image/png")
    return HTMLResponse(status_code=404)


@app.get("/api/status")
async def api_status():
    """Checks Tally Prime connection on Port 9000 and returns active open company."""
    status = tally_client.get_tally_status()
    return status


@app.get("/api/groups")
async def api_groups():
    """Fetches all ledger groups, subgroups, and hierarchy metadata."""
    return tally_client.get_groups_with_ledgers_and_meta()


@app.post("/api/vouchers")
async def api_vouchers(query: VoucherQuery):
    """Fetches vouchers from Tally Prime for the selected group, ledgers, and date range."""
    records = tally_client.fetch_vouchers_from_tally(
        from_date=query.from_date,
        to_date=query.to_date,
        group_name=query.group_name,
        selected_ledgers=query.selected_ledgers
    )
    return {
        "count": len(records),
        "vouchers": records
    }


@app.post("/api/export-excel")
async def api_export_excel(req: ExportRequest):
    """Generates and downloads the formatted 20-column Excel report with active formulas."""
    try:
        temp_dir = tempfile.gettempdir()
        clean_name = re.sub(r'[\/\\:*?"<>|\s]+', '_', req.party_name.strip()) or "Voucher_Report"
        filename = f"Tally_Voucher_Report_{clean_name}.xlsx"
        output_path = os.path.join(temp_dir, filename)

        report_engine.generate_voucher_excel(
            vouchers=req.vouchers,
            party_name=req.party_name,
            output_path=output_path
        )

        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/merge-excel")
async def api_merge_excel(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    merge_mode: str = Form("singleSheet"),
    add_source_col: bool = Form(False)
):
    """Merges multiple uploaded Excel files and returns the consolidated workbook."""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    temp_dir = tempfile.mkdtemp()
    saved_paths = []

    try:
        for file in files:
            safe_name = os.path.basename(file.filename or "upload.xlsx")
            dest_path = os.path.join(temp_dir, safe_name)
            with open(dest_path, "wb") as f:
                content = await file.read()
                f.write(content)
            saved_paths.append(dest_path)

        output_filename = "Consolidated_Tally_Merged_Report.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        report_engine.merge_excel_workbooks(
            file_paths=saved_paths,
            mode=merge_mode,
            add_source_col=add_source_col,
            output_path=output_path
        )

        # Cleanup temporary uploaded files and folder after streaming is finished
        background_tasks.add_task(shutil.rmtree, temp_dir, ignore_errors=True)

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))


def start_server(host="127.0.0.1", port=8000, open_browser=True):
    """Starts the uvicorn server and automatically opens the browser."""
    url = f"http://{host}:{port}"
    print(f"=======================================================")
    print(f"  TallyDataX Server running at: {url}")
    print(f"  Tally Prime Port: 9000")
    print(f"=======================================================")

    if open_browser:
        def launch_browser():
            import time
            time.sleep(1.0)
            webbrowser.open(url)
        threading.Thread(target=launch_browser, daemon=True).start()

    try:
        uvicorn.run(app, host=host, port=port, log_level="info")
    except OSError as e:
        if "10048" in str(e) or "address already in use" in str(e).lower():
            print(f"\n[INFO] TallyDataX server is already active on {url}!")
            print(f"[INFO] Opening existing instance in your browser...\n")
            if open_browser:
                webbrowser.open(url)
        else:
            raise e


if __name__ == "__main__":
    start_server(open_browser=True)
