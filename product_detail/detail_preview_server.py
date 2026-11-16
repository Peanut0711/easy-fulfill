"""상세 HTML을 이 PC에서만 제공하는 편집/브라우저 미리보기 서버."""

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import secrets
import threading
from urllib.parse import unquote, urlsplit


IMAGE_TYPES = {
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml",
    ".bmp": "image/bmp", ".avif": "image/avif",
}


class DetailPreviewServer:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self._prefix = f"/{secrets.token_urlsafe(24)}/"
        self._html = None
        self._revision = 0
        self.closed = False
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                self.respond(False)

            def do_HEAD(self):
                self.respond(True)

            def respond(self, head):
                # 다른 호스트 이름이나 임의 경로로 로컬 파일을 열지 않는다.
                if self.headers.get("Host") != owner._authority:
                    self.send_error(403)
                    return
                request_path = unquote(urlsplit(self.path).path)
                if not request_path.startswith(owner._prefix):
                    self.send_error(404)
                    return
                relative = request_path[len(owner._prefix):]
                try:
                    if relative == "preview.html":
                        content = owner._html
                        if content is None:
                            content = owner.path.read_bytes()
                        content_type = "text/html; charset=utf-8"
                    else:
                        asset = (owner.path.parent / relative).resolve()
                        content_type = IMAGE_TYPES.get(asset.suffix.lower())
                        if not content_type or not asset.is_relative_to(owner.path.parent) or not asset.is_file():
                            self.send_error(404)
                            return
                        content = asset.read_bytes()
                except (OSError, ValueError):
                    self.send_error(404)
                    return
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(content)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Referrer-Policy", "strict-origin-when-cross-origin")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.end_headers()
                if not head:
                    try:
                        self.wfile.write(content)
                    except (BrokenPipeError, ConnectionResetError):
                        pass  # 연속 편집으로 직전 요청이 취소될 수 있다.

            def log_message(self, *_args):
                pass

        self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self._authority = f"127.0.0.1:{self._server.server_port}"
        self.url = f"http://{self._authority}{self._prefix}preview.html"
        self._thread = threading.Thread(
            target=self._server.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True,
        )
        self._thread.start()

    def update(self, source):
        if self.closed:
            raise RuntimeError("미리보기 서버가 종료되었습니다.")
        self._html = source.encode("utf-8")
        self._revision += 1
        return f"{self.url}?revision={self._revision}"

    def close(self, *_args):
        if self.closed:
            return
        self.closed = True
        self._server.shutdown()
        self._server.server_close()
        self._thread.join(timeout=1)
