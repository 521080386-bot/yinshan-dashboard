#!/usr/bin/env python3
"""
截图保存服务 - 端口 5176
自动保存截图到桌面
"""
import base64, json
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

PORT = 5176
SAVE_DIR = Path.home() / "Desktop"

def cors_headers():
    return [
        ("Access-Control-Allow-Origin", "*"),
        ("Access-Control-Allow-Methods", "POST, OPTIONS"),
        ("Access-Control-Allow-Headers", "Content-Type"),
    ]

class Handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        for k, v in cors_headers():
            self.send_header(k, v)
        self.end_headers()

    def do_POST(self):
        if self.path == "/save-screenshot":
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                data = json.loads(body)
                img_b64 = data.get("image", "")
                if img_b64.startswith("data:image/png;base64,"):
                    img_b64 = img_b64.split(",", 1)[1]
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"ym_dashboard_{ts}.png"
                save_path = SAVE_DIR / filename
                save_path.write_bytes(base64.b64decode(img_b64))
                resp = {"ok": True, "path": str(save_path)}
                self.send_response(200)
                for k, v in cors_headers():
                    self.send_header(k, v)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(resp, ensure_ascii=False).encode())
                print(f"[save] {save_path}", flush=True)
            except Exception as e:
                self.send_response(500)
                for k, v in cors_headers():
                    self.send_header(k, v)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode())
        else:
            self.send_response(404)
            for k, v in cors_headers():
                self.send_header(k, v)
            self.end_headers()

    def log_message(self, fmt, *args):
        pass

if __name__ == "__main__":
    print(f"📸 截图保存服务 -> http://127.0.0.1:{PORT}/save-screenshot")
    HTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
