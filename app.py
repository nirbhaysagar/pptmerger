"""
app.py - High-Performance Web Server for PPT Unify
Handles batch uploading of up to 100 PPT/PPTX files error-free,
provides line-by-line ordering, sequential and interleaved merging,
and dual export to PPTX and PDF.
"""

import os
import sys
import json
import time
import uuid
import email
import zipfile
import logging
from email import policy
from urllib.parse import urlparse, parse_qs, unquote
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import threading

# Add current dir to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from merger_engine import PPTXMerger, PDFExporter, validate_pptx_file

logger = logging.getLogger("PPTUnifyApp")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# In-memory storage for active tasks & progress
MERGE_TASKS = {}


class PPTUnifyHandler(SimpleHTTPRequestHandler):
    """
    Custom HTTP Request Handler serving static frontend and API endpoints.
    """

    def log_message(self, format, *args):
        # Keep server console clean
        logger.info(f"{self.client_address[0]} - {args[0]}")

    def send_json(self, data, status_code=200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            self.serve_file(os.path.join(STATIC_DIR, "index.html"), "text/html; charset=utf-8")
        elif path.startswith("/static/"):
            rel_path = path[len("/static/"):]
            file_path = os.path.join(STATIC_DIR, rel_path)
            if os.path.exists(file_path):
                content_type = "text/plain"
                if file_path.endswith(".css"):
                    content_type = "text/css"
                elif file_path.endswith(".js"):
                    content_type = "application/javascript"
                elif file_path.endswith(".svg"):
                    content_type = "image/svg+xml"
                elif file_path.endswith(".png"):
                    content_type = "image/png"
                self.serve_file(file_path, content_type)
            else:
                self.send_error(404, "File Not Found")
        elif path == "/api/task-status":
            query = parse_qs(parsed.query)
            task_id = query.get("taskId", [""])[0]
            task = MERGE_TASKS.get(task_id, {"status": "unknown", "progress": 0, "message": "Task not found"})
            self.send_json(task)
        elif path.startswith("/api/download/"):
            filename = unquote(path[len("/api/download/"):])
            filepath = os.path.join(EXPORT_DIR, filename)
            if os.path.exists(filepath):
                self.serve_download(filepath)
            else:
                self.send_error(404, "Exported file not found.")
        else:
            self.send_error(404, "Not Found")

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/upload":
            self.handle_upload()
        elif path == "/api/merge":
            self.handle_merge()
        elif path == "/api/clear":
            self.handle_clear()
        else:
            self.send_error(404, "API endpoint not found")

    def serve_file(self, filepath, content_type):
        try:
            with open(filepath, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            logger.error(f"Error serving {filepath}: {e}")
            self.send_error(500, str(e))

    @staticmethod
    def _get_content_type(filename):
        if filename.endswith(".pptx"):
            return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif filename.endswith(".pdf"):
            return "application/pdf"
        elif filename.endswith(".zip"):
            return "application/zip"
        return "application/octet-stream"

    def do_HEAD(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path.startswith("/api/download/"):
            filename = unquote(path[len("/api/download/"):])
            filepath = os.path.join(EXPORT_DIR, filename)
            if os.path.exists(filepath):
                file_size = os.path.getsize(filepath)
                content_type = self._get_content_type(filename)
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(file_size))
                self.send_header("Accept-Ranges", "bytes")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Connection", "close")
                self.end_headers()
            else:
                self.send_error(404, "File not found")
        else:
            self.send_response(200)
            self.end_headers()

    def serve_download(self, filepath):
        try:
            filename = os.path.basename(filepath)
            file_size = os.path.getsize(filepath)
            content_type = self._get_content_type(filename)

            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(file_size))
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Connection", "close")
            self.end_headers()

            with open(filepath, "rb") as f:
                while True:
                    chunk = f.read(64 * 1024)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
            self.wfile.flush()
            logger.info(f"Successfully served download: {filename} ({file_size} bytes)")
        except (BrokenPipeError, ConnectionResetError):
            logger.debug(f"Client disconnected early during download of {filepath}")
        except Exception as e:
            logger.error(f"Error serving download {filepath}: {e}", exc_info=True)

    def handle_upload(self):
        """
        Receives uploaded files (handles multipart/form-data with up to 100 files).
        Validates each presentation and returns an error-free list of files.
        """
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            content_type = self.headers.get("Content-Type", "")

            if not content_type.startswith("multipart/form-data"):
                self.send_json({"success": False, "error": "Invalid Content-Type; expected multipart/form-data."}, 400)
                return

            # Read full body bytes
            body = self.rfile.read(content_length)

            # Parse multipart message using standard library email parser
            header_bytes = f"Content-Type: {content_type}\r\n\r\n".encode("utf-8")
            msg = email.message_from_bytes(header_bytes + body, policy=policy.default)

            uploaded_files = []

            for part in msg.iter_parts():
                filename = part.get_filename()
                if not filename:
                    continue

                # Clean filename
                clean_name = os.path.basename(filename)
                unique_prefix = uuid.uuid4().hex[:8]
                saved_filename = f"{unique_prefix}_{clean_name}"
                target_path = os.path.join(UPLOAD_DIR, saved_filename)

                # Save file data
                payload = part.get_payload(decode=True)
                with open(target_path, "wb") as f:
                    f.write(payload)

                # Validate PowerPoint file
                val = validate_pptx_file(target_path)
                uploaded_files.append({
                    "id": saved_filename,
                    "original_name": clean_name,
                    "size_bytes": val["file_size"],
                    "slide_count": val["slide_count"],
                    "is_valid": val["is_valid"],
                    "error_message": val["error_message"],
                    "aspect_ratio": val["aspect_ratio"],
                    "dimensions": f"{val['width_in']}\" x {val['height_in']}\""
                })

            self.send_json({
                "success": True,
                "files": uploaded_files,
                "count": len(uploaded_files)
            })

        except Exception as e:
            logger.error(f"Upload error: {e}", exc_info=True)
            self.send_json({"success": False, "error": f"Upload processing error: {str(e)}"}, 500)

    def handle_merge(self):
        """
        Starts an asynchronous merge job.
        Payload:
          - file_ids: list of filenames in user's requested order
          - mode: "sequential" or "interleaved"
          - export_format: "pptx", "pdf", or "both"
          - output_name: desired base name
        """
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length).decode("utf-8")
            data = json.loads(body)

            file_ids = data.get("file_ids", [])
            mode = data.get("mode", "sequential")
            export_format = data.get("export_format", "both")
            output_name = data.get("output_name", "Merged_Presentation").strip()
            if not output_name:
                output_name = "Merged_Presentation"

            # Sanitize output name
            output_name = "".join(c for c in output_name if c.isalnum() or c in ("-", "_", " ")).strip()
            if not output_name:
                output_name = "Merged_Presentation"

            if not file_ids:
                self.send_json({"success": False, "error": "No files selected to merge."}, 400)
                return

            task_id = uuid.uuid4().hex[:12]
            MERGE_TASKS[task_id] = {
                "status": "queued",
                "progress": 0,
                "message": "Initializing merge task...",
                "files_count": len(file_ids),
                "export_format": export_format,
                "downloads": {}
            }

            # Run in worker thread so HTTP request returns immediately with task_id
            def worker():
                run_merge_task(task_id, file_ids, mode, export_format, output_name)

            t = threading.Thread(target=worker, daemon=True)
            t.start()

            self.send_json({"success": True, "taskId": task_id})

        except Exception as e:
            logger.error(f"Merge trigger error: {e}", exc_info=True)
            self.send_json({"success": False, "error": str(e)}, 500)

    def handle_clear(self):
        """
        Cleans up uploads directory.
        """
        try:
            count = 0
            for f in os.listdir(UPLOAD_DIR):
                path = os.path.join(UPLOAD_DIR, f)
                if os.path.isfile(path):
                    os.remove(path)
                    count += 1
            self.send_json({"success": True, "cleared_count": count})
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, 500)


def run_merge_task(task_id, file_ids, mode, export_format, output_name):
    """
    Executes the PPT merge and PDF export in a background worker thread.
    """
    task = MERGE_TASKS[task_id]
    task["status"] = "processing"

    def update_progress(pct, msg):
        task["progress"] = pct
        task["message"] = msg
        logger.info(f"Task {task_id}: [{pct}%] {msg}")

    try:
        update_progress(5, "Preparing presentation files...")

        full_paths = [os.path.join(UPLOAD_DIR, fid) for fid in file_ids]
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        base_filename = f"{output_name}_{timestamp}"
        out_pptx_name = f"{base_filename}.pptx"
        out_pdf_name = f"{base_filename}.pdf"
        out_zip_name = f"{base_filename}.zip"

        out_pptx_path = os.path.join(EXPORT_DIR, out_pptx_name)
        out_pdf_path = os.path.join(EXPORT_DIR, out_pdf_name)
        out_zip_path = os.path.join(EXPORT_DIR, out_zip_name)

        # 1. Merge PPTX files
        merger = PPTXMerger(full_paths)
        merge_result = merger.merge(out_pptx_path, mode=mode, progress_callback=update_progress)

        task["downloads"]["pptx"] = f"/api/download/{out_pptx_name}"
        task["total_slides"] = merge_result["total_slides"]

        # 2. PDF Conversion if requested
        pdf_success = False
        if export_format in ("pdf", "both"):
            update_progress(90, "Generating PDF document...")
            pdf_success = PDFExporter.convert_to_pdf(out_pptx_path, out_pdf_path, progress_callback=update_progress)
            if pdf_success:
                task["downloads"]["pdf"] = f"/api/download/{out_pdf_name}"
            else:
                task["pdf_warning"] = "PDF conversion encountered an issue. PPTX is ready for download."

        # 3. Create ZIP package if both requested
        if export_format == "both":
            update_progress(99, "Packaging files...")
            with zipfile.ZipFile(out_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(out_pptx_path, arcname=f"{output_name}.pptx")
                if pdf_success and os.path.exists(out_pdf_path):
                    zf.write(out_pdf_path, arcname=f"{output_name}.pdf")
            task["downloads"]["zip"] = f"/api/download/{out_zip_name}"

        update_progress(100, "Done! Presentation merged successfully.")
        task["status"] = "completed"

    except Exception as e:
        logger.error(f"Merge task {task_id} failed: {e}", exc_info=True)
        task["status"] = "failed"
        task["error"] = str(e)
        task["message"] = f"Merge failed: {str(e)}"


def run_server(host="127.0.0.1", port=5005):
    """
    Starts the web server on specified or first available port.
    """
    actual_port = port
    server = None
    for p in range(port, port + 20):
        try:
            server = ThreadingHTTPServer((host, p), PPTUnifyHandler)
            actual_port = p
            break
        except OSError:
            continue

    if not server:
        raise RuntimeError("Could not find an available port to start server.")

    url = f"http://{host}:{actual_port}"
    print(f"\n=======================================================")
    print(f"🚀 PPT Unify is running at: {url}")
    print(f"📁 Upload up to 100 PPT/PPTX files, merge & export to PDF")
    print(f"=======================================================\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        server.server_close()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5005
    run_server(port=port)
