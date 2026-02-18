"""
Модуль перевірки оновлень для АЛЬВАРЕС AI.
Перевіряє GitHub Releases і повідомляє, якщо є нова версія.
"""

import urllib.request
import urllib.error
import json
from version import APP_VERSION, GITHUB_OWNER, GITHUB_REPO

_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
_RELEASES_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"


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

    Повертає словник {"version": str, "url": str, "notes": str}
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
            return {
                "version": latest_tag.lstrip("vV"),
                "url": data.get("html_url", _RELEASES_URL),
                "notes": data.get("body", "").strip(),
            }
    except Exception:
        pass

    return None


def get_releases_url() -> str:
    return _RELEASES_URL
