import os
import threading
import queue
import json
import io
import tempfile
import shutil
from datetime import datetime
from typing import List, Dict, Any

from flask import Flask, render_template, request, Response, jsonify, send_file
from werkzeug.utils import secure_filename

# Reuse existing extractor logic
from pdf_extractor_final import FinalPDFExtractor


app = Flask(__name__, template_folder="templates", static_folder="static")

# Configuration for file uploads
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# Simple in-memory event queue for SSE (single-user session)
event_queue: "queue.Queue[str]" = queue.Queue()

# Stores last processed table rows (so export can use edited rows from client)
last_rows_lock = threading.Lock()
last_rows: List[Dict[str, Any]] = []

# Progress snapshot for polling fallback
progress_lock = threading.Lock()
progress_state = {
    "statuses": [],  # list of {file, ok, items?, error?}
    "rows": [],      # final rows
    "running": False,
}


def sse_format(event: str, data: Any) -> str:
    payload = json.dumps(data, ensure_ascii=False)
    return f"event: {event}\ndata: {payload}\n\n"


def scan_and_extract(folder_path: str) -> None:
    """Worker: scan PDF folder, process each PDF, emit SSE events, and final table."""
    try:
        extractor = FinalPDFExtractor(input_folder=folder_path, output_folder="Python Output")
        pdf_files = extractor.list_pdf_files()

        if not pdf_files:
            event_queue.put(sse_format("error", {"message": "No PDF files found in the folder."}))
            event_queue.put(sse_format("complete", {"rows": []}))
            with progress_lock:
                progress_state["statuses"].append({"file": None, "ok": False, "error": "No PDF files found"})
                progress_state["rows"] = []
                progress_state["running"] = False
            return

        all_rows: List[Dict[str, Any]] = []

        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder_path, pdf_file)
            try:
                rows = extractor.extract_data_from_pdf(pdf_path)
                all_rows.extend(rows)
                status = {
                    "file": pdf_file,
                    "ok": len(rows) > 0,
                    "items": len(rows),
                }
                event_queue.put(sse_format("pdf_status", status))
                with progress_lock:
                    progress_state["statuses"].append(status)
            except Exception as e:  # pragma: no cover - defensive
                err = {"file": pdf_file, "ok": False, "error": str(e)}
                event_queue.put(sse_format("pdf_status", err))
                with progress_lock:
                    progress_state["statuses"].append(err)

        # Cache last rows for possible reuse
        with last_rows_lock:
            global last_rows
            last_rows = all_rows

        event_queue.put(sse_format("complete", {"rows": all_rows}))
        with progress_lock:
            progress_state["rows"] = all_rows
            progress_state["running"] = False
    except Exception as e:  # pragma: no cover - defensive
        event_queue.put(sse_format("error", {"message": str(e)}))
        event_queue.put(sse_format("complete", {"rows": []}))
        with progress_lock:
            progress_state["statuses"].append({"file": None, "ok": False, "error": str(e)})
            progress_state["rows"] = []
            progress_state["running"] = False


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/process")
def process_folder():
    # Check if we're in production mode (file upload) or development mode (folder path)
    is_production = os.getenv('ENV') == 'production'
    
    if is_production:
        # Production mode: handle file uploads
        return process_uploaded_files()
    else:
        # Development mode: handle folder path
        return process_folder_path()

def process_folder_path():
    """Process PDFs from a folder path (development mode)."""
    # Accept both JSON and form data for wizard compatibility
    if request.is_json:
        data = request.get_json(silent=True) or {}
        folder_path = data.get("folder_path", "").strip()
    else:
        folder_path = request.form.get("folder_path", "").strip()

    if not folder_path:
        return jsonify({"ok": False, "error": "Folder path is required."}), 400

    # Support remote/UNC paths as long as accessible to the server process
    if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
        return jsonify({"ok": False, "error": "Folder path does not exist or is not a directory."}), 400

    # Clear any stale events before starting new run
    try:
        while True:
            event_queue.get_nowait()
    except queue.Empty:
        pass

    # reset polling snapshot
    with progress_lock:
        progress_state["statuses"] = []
        progress_state["rows"] = []
        progress_state["running"] = True

    t = threading.Thread(target=scan_and_extract, args=(folder_path,), daemon=True)
    t.start()

    return jsonify({"ok": True})

def process_uploaded_files():
    """Process uploaded PDF files (production mode)."""
    if 'files' not in request.files:
        return jsonify({"ok": False, "error": "No files uploaded."}), 400
    
    files = request.files.getlist('files')
    
    if not files or all(file.filename == '' for file in files):
        return jsonify({"ok": False, "error": "No files selected."}), 400
    
    # Validate files
    valid_files = []
    for file in files:
        if file and allowed_file(file.filename):
            valid_files.append(file)
        else:
            return jsonify({"ok": False, "error": f"Invalid file: {file.filename}. Only PDF files are allowed."}), 400
    
    if not valid_files:
        return jsonify({"ok": False, "error": "No valid PDF files found."}), 400
    
    # Clear any stale events before starting new run
    try:
        while True:
            event_queue.get_nowait()
    except queue.Empty:
        pass

    # reset polling snapshot
    with progress_lock:
        progress_state["statuses"] = []
        progress_state["rows"] = []
        progress_state["running"] = True

    # Start processing in background thread
    t = threading.Thread(target=process_uploaded_files_worker, args=(valid_files,), daemon=True)
    t.start()

    return jsonify({"ok": True, "message": f"Processing {len(valid_files)} file(s)..."})

def process_uploaded_files_worker(files):
    """Worker thread to process uploaded files."""
    try:
        # Create temporary directory for uploaded files
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Save uploaded files to temp directory
            saved_files = []
            for file in files:
                filename = secure_filename(file.filename)
                file_path = os.path.join(temp_dir, filename)
                file.save(file_path)
                saved_files.append(file_path)
            
            # Process files using existing extractor logic
            extractor = FinalPDFExtractor(input_folder=temp_dir, output_folder="Python Output")
            all_rows: List[Dict[str, Any]] = []
            
            for file_path in saved_files:
                filename = os.path.basename(file_path)
                try:
                    rows = extractor.extract_data_from_pdf(file_path)
                    all_rows.extend(rows)
                    status = {
                        "file": filename,
                        "ok": len(rows) > 0,
                        "items": len(rows),
                    }
                    event_queue.put(sse_format("pdf_status", status))
                    with progress_lock:
                        progress_state["statuses"].append(status)
                except Exception as e:
                    err = {"file": filename, "ok": False, "error": str(e)}
                    event_queue.put(sse_format("pdf_status", err))
                    with progress_lock:
                        progress_state["statuses"].append(err)
            
            # Cache last rows for possible reuse
            with last_rows_lock:
                global last_rows
                last_rows = all_rows

            event_queue.put(sse_format("complete", {"rows": all_rows}))
            with progress_lock:
                progress_state["rows"] = all_rows
                progress_state["running"] = False
                
        finally:
            # Clean up temporary directory
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except Exception as e:
        event_queue.put(sse_format("error", {"message": str(e)}))
        event_queue.put(sse_format("complete", {"rows": []}))
        with progress_lock:
            progress_state["statuses"].append({"file": None, "ok": False, "error": str(e)})
            progress_state["rows"] = []
            progress_state["running"] = False


@app.get("/events")
def events():
    def stream():
        # initial hello
        yield sse_format("hello", {"message": "connected"})
        # heartbeat and event pump
        import time
        last_heartbeat = time.time()
        while True:
            try:
                # short timeout to interleave heartbeat
                data = event_queue.get(timeout=1.0)
                yield data
            except queue.Empty:
                pass
            # send heartbeat every 15s to keep connection alive
            now = time.time()
            if now - last_heartbeat >= 15:
                yield ":heartbeat\n\n"
                last_heartbeat = now

    headers = {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        # helpful for some proxies/servers
        "X-Accel-Buffering": "no",
    }
    return Response(stream(), headers=headers)


@app.get("/progress")
def get_progress():
    with progress_lock:
        snapshot = {
            "statuses": list(progress_state["statuses"]),
            "rows": list(progress_state["rows"]),
            "running": progress_state["running"],
        }
    return jsonify(snapshot)


@app.post("/showpath")
def showpath():
    # Legacy endpoint no longer needed in one-page flow; keep for compatibility
    folder_path = (request.form.get("folder_path") or "").strip()
    return render_template("index.html", folder_path=folder_path)


# Simple synchronous extraction for step-by-step flow
def extract_folder_sync(folder_path: str) -> Dict[str, Any]:
    extractor = FinalPDFExtractor(input_folder=folder_path, output_folder="Python Output")
    statuses: List[Dict[str, Any]] = []
    rows: List[Dict[str, Any]] = []
    pdf_files = extractor.list_pdf_files()
    for pdf in pdf_files:
        try:
            pdf_path = os.path.join(folder_path, pdf)
            data = extractor.extract_data_from_pdf(pdf_path)
            rows.extend(data)
            statuses.append({"file": pdf, "ok": len(data) > 0, "items": len(data)})
        except Exception as e:
            statuses.append({"file": pdf, "ok": False, "error": str(e)})
    return {"statuses": statuses, "rows": rows}


@app.post("/process_sync")
def process_sync():
    folder_path = (request.form.get("folder_path") or "").strip()
    if not folder_path:
        return render_template("index.html", folder_path=folder_path, error="Folder path required")
    if not os.path.isdir(folder_path):
        return render_template("index.html", folder_path=folder_path, error="Folder not found or not a directory")

    result = extract_folder_sync(folder_path)
    return render_template("index.html", folder_path=folder_path, statuses=result["statuses"], rows=result["rows"]) 


@app.post("/export")
def export_excel():
    """Export edited table rows to Excel and return as file download."""
    try:
        payload = request.get_json(force=True)
    except Exception:  # pragma: no cover - defensive
        payload = {}

    rows: List[Dict[str, Any]] = payload.get("rows") or []

    # If client didn't send rows, fall back to last processed rows
    if not rows:
        with last_rows_lock:
            rows = list(last_rows)

    if not rows:
        return jsonify({"ok": False, "error": "No rows to export."}), 400

    # Build Excel in-memory
    try:
        import pandas as pd

        df = pd.DataFrame(rows)

        # Preserve column order similar to extractor output if columns exist
        column_order = [
            "Vendor", "PO Number", "PO Date", "Vendor PO", "Customer Name", "Kama SKU",
            "Metal KT", "Metal Color", "Size", "Ship Date", "Qty",
            "Dia Quality", "ItemType", "Engraving", "Desc", "Category",
            "Line No", "Metal Tol.", "Dia Tol.", "Req Date",
        ]
        existing_columns = [c for c in column_order if c in df.columns]
        remaining_columns = [c for c in df.columns if c not in existing_columns]
        df = df[existing_columns + remaining_columns]

        output = io.BytesIO()
        # engine openpyxl available from requirements
        df.to_excel(output, index=False)
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"edited_data_{timestamp}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:  # pragma: no cover - defensive
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    # Development server - use Flask dev server
    # Bind to 0.0.0.0 so it can be accessed remotely if needed; enable threading for SSE
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)


