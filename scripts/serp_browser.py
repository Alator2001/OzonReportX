"""
Браузер для парсера выдачи: Playwright (Chromium) с профилем региона (storage_state) и необязательным прокси.

Playwright — необязательная зависимость: ставится только при включении парсера
(`python scripts/serp_profile.py setup`), поэтому импортируется внутри функций.
"""
import json
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, Optional
from urllib.parse import quote, unquote, urlsplit

from dotenv import load_dotenv

try:
    from scripts import serp_parser as parser
    from scripts import serp_paths as paths
except ModuleNotFoundError:
    import serp_parser as parser
    import serp_paths as paths

INSTALL_HINT = ("Парсер выдачи требует Playwright. Установите его один раз: "
                ".venv/bin/python scripts/serp_profile.py setup (на Windows — .venv\\Scripts\\python.exe).")


def playwright_available() -> bool:
    try:
        import playwright.sync_api  # noqa: F401
    except ImportError:
        return False
    return True


def proxy_settings(env_name: Optional[str]) -> Optional[Dict[str, str]]:
    """Прокси региона из .env по имени переменной: http://user:pass@host:port → настройки Playwright."""
    if not env_name:
        return None
    load_dotenv()
    url = (os.getenv(env_name) or "").strip()
    if not url:
        raise RuntimeError(f"В .env нет переменной {env_name} с адресом прокси для региона.")
    parts = urlsplit(url)
    if not parts.hostname:
        raise RuntimeError(f"{env_name}: не удалось разобрать адрес прокси — нужен вид http://логин:пароль@хост:порт.")
    proxy = {"server": f"{parts.scheme or 'http'}://{parts.hostname}" + (f":{parts.port}" if parts.port else "")}
    if parts.username:
        proxy["username"] = unquote(parts.username)
    if parts.password:
        proxy["password"] = unquote(parts.password)
    return proxy


@contextmanager
def open_page(profile_path: Optional[Path], proxy_env: Optional[str], headless: bool, timeout_seconds: float) -> Iterator[Any]:
    """Страница Chromium с профилем региона. На выходе браузер закрывается в любом случае."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless, proxy=proxy_settings(proxy_env))
        try:
            context = browser.new_context(
                storage_state=str(profile_path) if profile_path and Path(profile_path).exists() else None,
                locale="ru-RU", timezone_id="Europe/Moscow", viewport={"width": 1440, "height": 900},
            )
            context.set_default_timeout(timeout_seconds * 1000)
            yield context.new_page()
        finally:
            browser.close()


def open_site(page) -> None:
    """Главная Ozon — как обычный заход на сайт. Капча на главной — SerpBlocked."""
    response = page.goto(paths.SITE + "/", wait_until="domcontentloaded")
    title = page.title()
    if (response is not None and response.status in (403, 429)) or "antibot" in title.lower():
        raise parser.SerpBlocked(f"Ozon показал капчу на главной странице (HTTP {response.status if response else '—'}).")


def fetch_page_json(page, site_path: str) -> Dict[str, Any]:
    """Порция выдачи через JSON-эндпоинт тем же запросом, что делает сам сайт (с cookies профиля)."""
    result = page.evaluate("""async (url) => {
        const r = await fetch(url, {credentials: 'include', headers: {'Accept': 'application/json'}});
        return {status: r.status, contentType: r.headers.get('content-type') || '', text: await r.text()};
    }""", paths.JSON_ENDPOINT + quote(site_path, safe=""))
    return {"data": parser.check_response(result["status"], result["contentType"], result["text"]), "raw": result["text"]}


def save_state(page, profile_path: Path) -> Dict[str, Any]:
    """Сохраняет cookies и localStorage профиля (Ozon обновляет их при каждом заходе)."""
    Path(profile_path).parent.mkdir(parents=True, exist_ok=True)
    state = page.context.storage_state()
    tmp = Path(profile_path).with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, profile_path)
    return state
