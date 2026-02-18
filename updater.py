"""
Модуль перевірки оновлень для АЛЬВАРЕС AI.
Перевіряє GitHub Releases і повідомляє, якщо є нова версія.
Підтримує завантаження та тиху установку оновлень.
"""

import os
import sys
import subprocess
import tempfile
import urllib.request
import urllib.error
import json
from version import APP_VERSION, GITHUB_OWNER, GITHUB_REPO

_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
_RELEASES_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
_SETUP_FILENAME = "AlvaresAI_Setup.exe"


def _parse_version(v: str) -> tuple:
    """Перетворює рядок версії типу '1.2.3' у кортеж (1, 2, 3)."""
    v = v.lstrip("vV").strip()
    try:
        return tuple(int(x) for x in v.split("."))
    except ValueError:
        return (0,)


def check_for_update(timeout: int = 5) -> dict | None:
    """
    Перевіряє GitHub Releases на наявність нової версії.

    Повертає словник {"version": str, "url": str, "notes": str, "download_url": str | None}
    якщо нова версія доступна, або None якщо версія актуальна або
    сталася помилка мережі.
    """
    try:
        req = urllib.request.Request(
            _API_URL,
            headers={"User-Agent": f"AlvaresAI/{APP_VERSION}"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode())

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None

        if _parse_version(latest_tag) > _parse_version(APP_VERSION):
            download_url = None
            for asset in data.get("assets", []):
                name = asset.get("name", "")
                if name.lower().endswith(".exe"):
                    download_url = asset.get("browser_download_url")
                    break

            return {
                "version": latest_tag.lstrip("vV"),
                "url": data.get("html_url", _RELEASES_URL),
                "notes": data.get("body", "").strip(),
                "download_url": download_url,
            }
    except Exception:
        pass

    return None


def download_update(download_url: str, dest_path: str | None = None,
                    on_progress=None) -> str:
    """
    Завантажує Setup.exe з GitHub Release.

    Args:
        download_url: URL .exe файлу з GitHub Release
        dest_path: шлях для збереження (якщо None — temp директорія)
        on_progress: callback(downloaded_bytes, total_bytes) для прогресу

    Returns:
        Шлях до завантаженого файлу

    Raises:
        Exception: якщо завантаження не вдалося
    """
    if dest_path is None:
        dest_path = os.path.join(tempfile.gettempdir(), _SETUP_FILENAME)

    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": f"AlvaresAI/{APP_VERSION}"}
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        chunk_size = 64 * 1024

        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if on_progress:
                    on_progress(downloaded, total)

    return dest_path


def install_update(setup_path: str):
    """
    Запускає інсталятор у тихому режимі і закриває поточний додаток.
    Inno Setup оновить файли і запустить нову версію (секція [Run]).
    """
    subprocess.Popen(
        [setup_path, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    sys.exit(0)


def get_releases_url() -> str:
    return _RELEASES_URL
