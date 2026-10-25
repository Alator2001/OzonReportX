"""
Профили регионов для парсера выдачи. Запускать из терминала: создание профиля требует выбрать город руками.

    python scripts/serp_profile.py setup                                -> установить Playwright и Chromium (~150 МБ)
    python scripts/serp_profile.py create msk "Москва"                  -> открыть браузер, выбрать город, сохранить профиль
    python scripts/serp_profile.py create spb "Санкт-Петербург" --proxy-env SERP_PROXY_SPB
    python scripts/serp_profile.py list
    python scripts/serp_profile.py disable spb | enable spb
    python scripts/serp_profile.py diff msk spb                         -> чем отличаются cookies двух профилей

Профиль — файл storage_state Playwright (cookies и localStorage анонимной сессии Ozon) в `search data/serp_profiles/`.
Код ФИАС города берётся из ответа Ozon после выбора города и потом проверяется перед каждым замером.
"""
import argparse
import hashlib
import json
import re
import subprocess
import sys
from contextlib import closing
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from scripts import search_db as db
    from scripts import serp_browser as browser
    from scripts import serp_parser as parser
except ModuleNotFoundError:
    import search_db as db
    import serp_browser as browser
    import serp_parser as parser

PROFILES_DIR = db.DATA_DIR / "serp_profiles"
CODE_RE = re.compile(r"^[a-z0-9_-]{2,32}$")


def profile_path(code: str) -> Path:
    return PROFILES_DIR / f"{code}.json"


def cookie_fingerprint(state: Dict[str, Any]) -> Dict[str, str]:
    """Имена cookies → хеш значения: сами значения (токены сессии) в отчёт не попадают."""
    return {f"{c.get('domain')}|{c.get('name')}": hashlib.sha1(str(c.get("value")).encode()).hexdigest()[:10]
            for c in state.get("cookies") or []}


def diff_fingerprints(a: Dict[str, str], b: Dict[str, str]) -> Dict[str, List[str]]:
    return {"only_a": sorted(set(a) - set(b)), "only_b": sorted(set(b) - set(a)),
            "different": sorted(k for k in set(a) & set(b) if a[k] != b[k])}


def setup() -> int:
    print("Устанавливаю Playwright и Chromium (около 150 МБ)…", flush=True)
    for cmd in ([sys.executable, "-m", "pip", "install", "playwright"], [sys.executable, "-m", "playwright", "install", "chromium"]):
        code = subprocess.call(cmd)
        if code:
            print(f"❌ Команда завершилась с кодом {code}: {' '.join(cmd)}")
            return code
    print("✅ Готово. Теперь создайте профили регионов: scripts/serp_profile.py create msk \"Москва\"")
    return 0


def create(code: str, name: str, proxy_env: Optional[str]) -> int:
    if not CODE_RE.match(code):
        print("Код региона — латиница, цифры, «-» или «_», 2–32 символа (например, msk, spb).")
        return 1
    if not browser.playwright_available():
        print(browser.INSTALL_HINT)
        return 1
    path = profile_path(code)
    with browser.open_page(None, proxy_env, headless=False, timeout_seconds=60) as page:
        browser.open_site(page)
        print(f"\nВ открывшемся окне выберите город «{name}» в шапке Ozon (адрес указывать не нужно).")
        input("Когда город выбран и страница обновилась, нажмите Enter здесь… ")
        response = browser.fetch_page_json(page, "/")
        got = parser.location(response["data"])
        if not got["fias"]:
            print("❌ Ozon не сообщил регион страницы — профиль не сохранён. Попробуйте ещё раз.")
            return 1
        print(f"Ozon считает регионом: {got['city']} (ФИАС {got['fias']}).")
        if (got["city"] or "").strip().lower() != name.strip().lower():
            answer = input(f"Это не «{name}». Всё равно сохранить профиль? [y/N] ").strip().lower()
            if answer not in ("y", "yes", "д", "да"):
                print("Профиль не сохранён.")
                return 1
        state = browser.save_state(page, path)
    fingerprint_path = path.with_suffix(".cookies.json")
    fingerprint_path.write_text(json.dumps(cookie_fingerprint(state), ensure_ascii=False, indent=1), encoding="utf-8")
    with closing(db.connect()) as conn, conn:
        db.upsert_region(conn, code, name, fias=got["fias"], profile_path=str(path.relative_to(db.ROOT)),
                         proxy_env=proxy_env, active=1)
    print(f"✅ Профиль «{name}» сохранён: {path.relative_to(db.ROOT)}")
    return 0


def list_regions() -> int:
    with closing(db.connect()) as conn:
        rows = conn.execute("SELECT * FROM regions ORDER BY id").fetchall()
    if not rows:
        print("Регионов пока нет. Создайте: scripts/serp_profile.py create msk \"Москва\"")
    for r in rows:
        exists = r["profile_path"] and (db.ROOT / r["profile_path"]).exists()
        print(f"{r['code']:8} {r['name']:20} {'активен' if r['active'] else 'выключен':9} ФИАС {r['fias'] or '—'}  "
              f"профиль {'есть' if exists else 'НЕТ'}  прокси {r['proxy_env'] or '—'}")
    return 0


def set_active(code: str, active: bool) -> int:
    with closing(db.connect()) as conn, conn:
        changed = conn.execute("UPDATE regions SET active = ?, updated_at = ? WHERE code = ?", (int(active), db.now_iso(), code)).rowcount
    print("Готово." if changed else f"Региона «{code}» нет.")
    return 0 if changed else 1


def diff(code_a: str, code_b: str) -> int:
    prints = {}
    for code in (code_a, code_b):
        path = profile_path(code).with_suffix(".cookies.json")
        if not path.exists():
            print(f"Нет снимка cookies профиля «{code}» — создайте профиль заново.")
            return 1
        prints[code] = json.loads(path.read_text(encoding="utf-8"))
    result = diff_fingerprints(prints[code_a], prints[code_b])
    print(f"Только в {code_a}: {', '.join(result['only_a']) or '—'}")
    print(f"Только в {code_b}: {', '.join(result['only_b']) or '—'}")
    print(f"Разные значения: {', '.join(result['different']) or '—'}")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    args_parser = argparse.ArgumentParser(description="Профили регионов для парсера выдачи Ozon.")
    sub = args_parser.add_subparsers(dest="command", required=True)
    sub.add_parser("setup")
    c = sub.add_parser("create")
    c.add_argument("code")
    c.add_argument("name")
    c.add_argument("--proxy-env", help="имя переменной в .env с адресом прокси для этого региона")
    sub.add_parser("list")
    for name in ("enable", "disable"):
        sub.add_parser(name).add_argument("code")
    d = sub.add_parser("diff")
    d.add_argument("code_a")
    d.add_argument("code_b")
    args = args_parser.parse_args(argv)
    if args.command == "setup":
        return setup()
    if args.command == "create":
        return create(args.code, args.name, args.proxy_env)
    if args.command == "list":
        return list_regions()
    if args.command in ("enable", "disable"):
        return set_active(args.code, args.command == "enable")
    return diff(args.code_a, args.code_b)


if __name__ == "__main__":
    sys.exit(main())
