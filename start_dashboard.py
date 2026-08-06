import http.server
import hashlib
import os
import socket
import socketserver
import webbrowser

DEFAULT_PORT = 8000
BUILD_TAG = "20260806a"


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()


def find_open_port(start_port: int) -> int:
    port = start_port
    while port < start_port + 100:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            if probe.connect_ex(("127.0.0.1", port)) != 0:
                return port
        port += 1
    return start_port


def file_sha256(path: str) -> str:
    if not os.path.exists(path):
        return "MISSING"

    digest = hashlib.sha256()
    with open(path, "rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()[:12]


def main() -> None:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    port = find_open_port(DEFAULT_PORT)
    url = f"http://localhost:{port}/?v={BUILD_TAG}"

    index_path = os.path.join(base_dir, "index.html")
    app_path = os.path.join(base_dir, "app.js")
    index_hash = file_sha256(index_path)
    app_hash = file_sha256(app_path)

    with socketserver.TCPServer(("", port), NoCacheHandler) as httpd:
        print(f"Dashboard build: {BUILD_TAG}")
        print(f"Launcher path: {os.path.abspath(__file__)}")
        print(f"Working dir: {os.getcwd()}")
        print(f"index.html SHA256: {index_hash}")
        print(f"app.js SHA256: {app_hash}")
        print(f"Serving project dashboard at {url}")
        webbrowser.open(url)
        httpd.serve_forever()


if __name__ == "__main__":
    main()
