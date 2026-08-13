import hashlib
import http.server
import json
import os
import socket
import socketserver
import tempfile
from urllib.parse import urlsplit
import webbrowser

DEFAULT_PORT = 8000
BUILD_TAG = "20260813a"
ROOT_FOLDERS = {"EI", "SCP", "SER", "WO"}
COMPLETED_FOLDERS = {"complete", "completed"}
MAX_REQUEST_BYTES = 16 * 1024


class ProjectOperationError(Exception):
    def __init__(self, status_code: int, message: str) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.message = message


def get_projects_root(base_dir: str) -> str:
    configured_root = os.environ.get("PROJECTS_ROOT", "").strip()
    if configured_root:
        return os.path.abspath(configured_root)

    return os.path.abspath(os.path.join(base_dir, os.pardir))


def validate_path_component(value: object, label: str, allowed_values=None) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ProjectOperationError(400, label + " is required.")

    component = value.strip()
    invalid_characters = set('\\/:*?"<>|')
    if (
        component in {".", ".."}
        or "\x00" in component
        or any(character in invalid_characters for character in component)
    ):
        raise ProjectOperationError(400, "Invalid " + label + ".")

    if allowed_values is not None and component not in allowed_values:
        raise ProjectOperationError(400, "Unsupported project type.")

    return component


def find_completed_directory(building_dir: str) -> str:
    try:
        entries = sorted(os.scandir(building_dir), key=lambda entry: entry.name.casefold())
    except OSError as error:
        raise ProjectOperationError(500, "Unable to inspect the building folder: " + str(error)) from error

    for entry in entries:
        if entry.is_dir(follow_symlinks=False) and entry.name.casefold() in COMPLETED_FOLDERS:
            return entry.path

    completed_dir = os.path.join(building_dir, "Completed")
    try:
        os.mkdir(completed_dir)
    except FileExistsError:
        if not os.path.isdir(completed_dir):
            raise ProjectOperationError(409, "A file named Completed already exists.")
    except OSError as error:
        raise ProjectOperationError(500, "Unable to create the Completed folder: " + str(error)) from error

    return completed_dir


def read_project_data(project_json_path: str) -> dict:
    try:
        with open(project_json_path, "r", encoding="utf-8-sig") as file_handle:
            data = json.load(file_handle)
    except (OSError, json.JSONDecodeError) as error:
        raise ProjectOperationError(409, "The project.json file could not be read as valid JSON.") from error

    if not isinstance(data, dict):
        raise ProjectOperationError(409, "The project.json root must be an object.")

    return data


def set_completed_metadata(data: dict) -> None:
    data["percentComplete"] = 100
    if "status" in data:
        data["status"] = "Closed"
    if "Status" in data:
        data["Status"] = "Closed"
    if "status" not in data and "Status" not in data:
        data["status"] = "Closed"


def write_project_data(project_json_path: str, data: dict) -> None:
    directory = os.path.dirname(project_json_path)
    temporary_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=directory,
            prefix=".project-dashboard-",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = temporary_file.name
            json.dump(data, temporary_file, indent="\t")
            temporary_file.write("\n")
        os.replace(temporary_path, project_json_path)
    except OSError as error:
        if temporary_path:
            try:
                os.remove(temporary_path)
            except OSError:
                pass
        raise ProjectOperationError(500, "Unable to update project.json: " + str(error)) from error


def complete_project(payload: object, projects_root: str) -> dict:
    if not isinstance(payload, dict):
        raise ProjectOperationError(400, "The request body must be a JSON object.")

    root_name = validate_path_component(payload.get("rootName"), "rootName", ROOT_FOLDERS)
    building_name = validate_path_component(payload.get("buildingName"), "buildingName")
    folder_name = validate_path_component(payload.get("folderName"), "folderName")
    if folder_name.casefold() in COMPLETED_FOLDERS:
        raise ProjectOperationError(400, "A completed folder cannot be completed again.")

    allowed_root = os.path.realpath(projects_root)
    root_dir = os.path.abspath(os.path.join(projects_root, root_name))
    building_dir = os.path.abspath(os.path.join(root_dir, building_name))
    source_dir = os.path.abspath(os.path.join(building_dir, folder_name))
    source_real_path = os.path.realpath(source_dir)

    try:
        if os.path.commonpath([allowed_root, source_real_path]) != allowed_root:
            raise ProjectOperationError(400, "The project path is outside the configured Projects folder.")
    except ValueError as error:
        raise ProjectOperationError(400, "The project path is outside the configured Projects folder.") from error

    if not os.path.isdir(building_dir):
        raise ProjectOperationError(404, "The building folder was not found.")
    if not os.path.isdir(source_dir):
        raise ProjectOperationError(404, "The active project folder was not found.")
    if os.path.islink(source_dir):
        raise ProjectOperationError(409, "The active project folder cannot be a link.")

    project_json_path = os.path.join(source_dir, "project.json")
    if not os.path.isfile(project_json_path):
        raise ProjectOperationError(409, "The project folder does not contain project.json.")

    project_data = read_project_data(project_json_path)
    set_completed_metadata(project_data)
    completed_dir = find_completed_directory(building_dir)
    destination_dir = os.path.abspath(os.path.join(completed_dir, folder_name))
    if os.path.lexists(destination_dir):
        raise ProjectOperationError(409, "A completed project with this folder name already exists.")

    try:
        os.rename(source_dir, destination_dir)
    except OSError as error:
        raise ProjectOperationError(
            409,
            "Windows could not move the project folder. It may be open or unavailable: " + str(error),
        ) from error

    destination_json_path = os.path.join(destination_dir, "project.json")
    try:
        write_project_data(destination_json_path, project_data)
    except ProjectOperationError:
        try:
            os.rename(destination_dir, source_dir)
        except OSError:
            pass
        raise

    completed_folder_name = os.path.basename(completed_dir)
    return {
        "ok": True,
        "rootName": root_name,
        "buildingName": building_name,
        "folderName": folder_name,
        "completedFolderName": completed_folder_name,
        "projectPath": os.path.join(root_name, building_name, completed_folder_name, folder_name),
    }


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self) -> None:
        if urlsplit(self.path).path != "/api/complete-project":
            self.send_json(404, {"ok": False, "error": "Endpoint not found."})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json(400, {"ok": False, "error": "Invalid request length."})
            return

        if content_length <= 0 or content_length > MAX_REQUEST_BYTES:
            self.send_json(413, {"ok": False, "error": "Request body is missing or too large."})
            return

        try:
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            result = complete_project(payload, get_projects_root(os.path.dirname(os.path.abspath(__file__))))
        except json.JSONDecodeError:
            self.send_json(400, {"ok": False, "error": "Request body must be valid JSON."})
        except ProjectOperationError as error:
            self.send_json(error.status_code, {"ok": False, "error": error.message})
        except (UnicodeDecodeError, OSError) as error:
            self.send_json(500, {"ok": False, "error": "Unable to complete the project: " + str(error)})
        else:
            self.send_json(200, result)

    def send_json(self, status_code: int, payload: dict) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

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
    projects_root = get_projects_root(base_dir)

    with socketserver.TCPServer(("127.0.0.1", port), NoCacheHandler) as httpd:
        print(f"Dashboard build: {BUILD_TAG}")
        print(f"Launcher path: {os.path.abspath(__file__)}")
        print(f"Working dir: {os.getcwd()}")
        print(f"index.html SHA256: {index_hash}")
        print(f"app.js SHA256: {app_hash}")
        print(f"Projects root: {projects_root}")
        print(f"Serving project dashboard at {url}")
        webbrowser.open(url)
        httpd.serve_forever()


if __name__ == "__main__":
    main()