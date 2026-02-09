import http.server
import os
import socket
import socketserver
import webbrowser

DEFAULT_PORT = 8000


def find_open_port(start_port: int) -> int:
    port = start_port
    while port < start_port + 100:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            if probe.connect_ex(("127.0.0.1", port)) != 0:
                return port
        port += 1
    return start_port


def main() -> None:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    port = find_open_port(DEFAULT_PORT)
    url = f"http://localhost:{port}/"

    handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(("", port), handler) as httpd:
        print(f"Serving project dashboard at {url}")
        webbrowser.open(url)
        httpd.serve_forever()


if __name__ == "__main__":
    main()
