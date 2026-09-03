#!/usr/bin/env python3
"""
Скрипт для создания подстраниц-паспортов систем в Яндекс Вики из Excel (XLSX).
"""

import os
import sys
import json
import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
import requests
import webbrowser
import urllib.parse
import http.server
import socketserver
import threading
import time
from http import HTTPStatus

# ========== НАСТРОЙКИ (ЗАМЕНИТЕ НА СВОИ) ==========
CLIENT_ID = ""
CLIENT_SECRET = ""
REDIRECT_URI = "https://oauth.yandex.ru/verification_code"
ORG_ID = os.environ.get("ORG_ID", "")  # ID организации

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PARENT_SLUG = ""
DEFAULT_DATA_FILE = os.path.join(SCRIPT_DIR, "")

# Поля содержимого страницы (порядок как в шаблоне вики)
PASSPORT_CONTENT_FIELDS = [
    "название системы",
    "Описание",
    "Заказчик",
    "Задача в трекере",
    "Тип системы",
    "Статус",
    "IP адрес",
    "Hostname",
    "Открытые порты",
]

# Подписи полей на странице вики
PASSPORT_FIELD_LABELS = {
    "название системы": "Название системы",
    "Описание": "Описание",
    "Заказчик": "Заказчик",
    "Задача в трекере": "Связанные задачи",
    "Тип системы": "Тип системы",
    "Статус": "Статус",
    "IP адрес": "IP адрес",
    "Hostname": "Hostname",
    "Открытые порты": "Открытые порты",
}

# Алиасы столбцов в файле (регистр не важен)
PASSPORT_FIELD_ALIASES = {
    "название системы": ["название системы"],
    "Описание": ["описание"],
    "Заказчик": ["заказчик"],
    "Задача в трекере": ["задача в трекере"],
    "Тип системы": ["тип системы"],
    "Статус": ["статус"],
    "IP адрес": ["ip адрес", "ip"],
    "Hostname": ["hostname", "scan_hostname"],
    "Открытые порты": ["открытые порты", "open_ports"],
}
# ==================================================

auth_code = None
server = None

# Сессия без системного прокси (часто ломает OAuth и API-запросы)
session = requests.Session()
session.trust_env = False


class OAuthHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/callback":
            query = urllib.parse.parse_qs(parsed.query)
            code = query.get("code")
            if code:
                auth_code = code[0]
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                html = "<h1>Авторизация успешна!</h1><p>Можете закрыть это окно.</p>"
                self.wfile.write(html.encode('utf-8'))
                threading.Timer(0.5, self.server.shutdown).start()
            else:
                self.send_response(HTTPStatus.BAD_REQUEST)
                self.end_headers()
                # Используем латиницу или кодируем
                self.wfile.write("Не удалось получить код.".encode('utf-8'))
        else:
            self.send_response(HTTPStatus.NOT_FOUND)
            self.end_headers()
            self.wfile.write(b"Unknown path.")


def start_server(port):
    global server
    server = socketserver.TCPServer(("localhost", port), OAuthHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()
    return thread


def stop_server():
    if server:
        server.shutdown()
        server.server_close()


def is_local_redirect(uri):
    parsed = urllib.parse.urlparse(uri)
    host = parsed.hostname
    return host in ("localhost", "127.0.0.1")


def get_oauth_token(client_id, client_secret, redirect_uri):
    global auth_code

    if is_local_redirect(redirect_uri):
        port = urllib.parse.urlparse(redirect_uri).port or 8080
        start_server(port)
        time.sleep(1)

        auth_url = (
            "https://oauth.yandex.ru/authorize"
            f"?response_type=code&client_id={client_id}&redirect_uri={redirect_uri}"
        )
        print("🌐 Открываю страницу авторизации в браузере...")
        webbrowser.open(auth_url)

        timeout = 120
        start_time = time.time()
        while auth_code is None and (time.time() - start_time) < timeout:
            time.sleep(0.5)

        stop_server()

        if auth_code is None:
            print("❌ Время ожидания истекло.")
            sys.exit(1)
        code = auth_code
    else:
        auth_url = (
            "https://oauth.yandex.ru/authorize"
            f"?response_type=code&client_id={client_id}&redirect_uri={redirect_uri}"
        )
        print("\n🔐 Перейдите по ссылке для авторизации:")
        print(auth_url)
        print("\nПосле авторизации скопируйте параметр 'code' из URL и вставьте его ниже.")
        code = input("Введите код авторизации: ").strip()
        if not code:
            print("❌ Код не введён.")
            sys.exit(1)

    print("✅ Код авторизации получен.")

    token_url = "https://oauth.yandex.ru/token"
    data = {
        "grant_type": "authorization_code",
        "code": code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
    }
    try:
        resp = session.post(token_url, data=data, timeout=30)
        resp.raise_for_status()
        token_data = resp.json()
        access_token = token_data.get("access_token")
        if not access_token:
            print("❌ Не удалось получить access_token. Ответ:", token_data)
            sys.exit(1)
        print("✅ Access_token получен.")
        return access_token
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при обмене кода на токен: {e}")
        if hasattr(e, 'response') and e.response:
            print(e.response.text)
        sys.exit(1)


def normalize_slug(value: str) -> str:
    """Извлекает чистый slug из URL, пути или slug с query-параметрами."""
    value = value.strip()
    if not value:
        return value

    if not value.startswith("http") and "wiki.yandex.ru" in value:
        value = "https://" + value.lstrip("/")

    if value.startswith("http"):
        parsed = urllib.parse.urlparse(value)
        path = parsed.path
    else:
        parsed = urllib.parse.urlparse(value)
        path = parsed.path if (parsed.query or parsed.fragment) else value

    path = path.strip("/")
    if path.startswith("wiki.yandex.ru/"):
        path = path[len("wiki.yandex.ru/"):]

    return path.strip("/")


def extract_slug_from_url(url: str) -> str:
    """Извлекает slug из полного URL Яндекс.Вики."""
    return normalize_slug(url)


def get_headers(token):
    return {
        "Host": "api.wiki.yandex.net",
        "Authorization": f"OAuth {token}",
        "X-Org-Id": ORG_ID,
        "Content-Type": "application/json",
    }


def get_page(token, slug, fields=None):
    """Получить информацию о странице."""
    url = "https://api.wiki.yandex.net/v1/pages"
    headers = get_headers(token)
    params = {"slug": slug}
    if fields:
        params["fields"] = fields
    try:
        resp = session.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при получении страницы {slug}: {e}", file=sys.stderr)
        if hasattr(e, 'response') and e.response:
            print(f"   Ответ: {e.response.text}", file=sys.stderr)
        return None


def get_page_by_id(token, page_id, fields=None):
    """Получить информацию о странице по ID."""
    url = f"https://api.wiki.yandex.net/v1/pages/{page_id}"
    headers = get_headers(token)
    params = {}
    if fields:
        params["fields"] = fields
    try:
        resp = session.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при получении страницы ID={page_id}: {e}", file=sys.stderr)
        if hasattr(e, 'response') and e.response:
            print(f"   Ответ: {e.response.text}", file=sys.stderr)
        return None


def update_page_content(token, page_id, content, title=None):
    """Обновить содержимое страницы."""
    url = f"https://api.wiki.yandex.net/v1/pages/{page_id}"
    headers = get_headers(token)
    body = {"content": content}
    if title:
        body["title"] = title
    try:
        resp = session.post(url, headers=headers, json=body, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при записи страницы ID={page_id}: {e}", file=sys.stderr)
        if hasattr(e, 'response') and e.response:
            print(f"   Ответ: {e.response.text}", file=sys.stderr)
        return None


def create_wiki_page(token, slug, title, content):
    """Создать новую страницу в вики."""
    url = "https://api.wiki.yandex.net/v1/pages"
    headers = get_headers(token)
    body = {
        "slug": slug,
        "title": title,
        "content": content,
    }
    try:
        resp = session.post(url, headers=headers, json=body, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при создании страницы {slug}: {e}", file=sys.stderr)
        if hasattr(e, 'response') and e.response:
            print(f"   Ответ: {e.response.text}", file=sys.stderr)
        return None


def slugify(text, max_len=80):
    """Преобразовать строку в slug для вики."""
    text = str(text or "").strip().lower()
    text = text.replace(".", "-")
    text = re.sub(r"[^\w\-]+", "-", text, flags=re.UNICODE)
    text = re.sub(r"-+", "-", text).strip("-")
    return text[:max_len] or "system"


def record_field(record, field_name):
    """Получить значение поля из строки файла (без учёта регистра)."""
    aliases = PASSPORT_FIELD_ALIASES.get(field_name, [field_name.lower()])
    lower_map = {str(k).lower().strip(): str(v or "").strip() for k, v in record.items()}
    for alias in aliases:
        if alias.lower() in lower_map:
            return lower_map[alias.lower()]
    return ""


def _ods_cell_text(cell):
    """Текст ячейки, включая гиперссылки."""
    return "".join(cell.itertext()).strip()


def read_ods(path):
    """Прочитать ODS-файл и вернуть список словарей."""
    if not os.path.isfile(path):
        print(f"❌ Файл не найден: {path}", file=sys.stderr)
        return None

    table_ns = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
    repeat_attr = f"{table_ns}number-columns-repeated"

    try:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("content.xml"))
            for table in root.findall(f".//{table_ns}table"):
                headers = None
                records = []
                for row in table.findall(f"{table_ns}table-row"):
                    cells = []
                    for cell in row.findall(f"{table_ns}table-cell"):
                        repeat = int(cell.get(repeat_attr) or 1)
                        value = _ods_cell_text(cell)
                        cells.extend([value] * repeat)
                    while cells and cells[-1] == "":
                        cells.pop()
                    if not cells:
                        continue
                    if headers is None:
                        headers = cells
                        continue
                    record = {
                        headers[i]: cells[i] if i < len(cells) else ""
                        for i in range(len(headers))
                    }
                    if any(str(v).strip() for v in record.values()):
                        records.append(record)
                if headers is not None:
                    return records
    except (OSError, ET.ParseError, KeyError, ValueError) as exc:
        print(f"❌ Ошибка чтения ODS {path}: {exc}", file=sys.stderr)
        return None

    return []


def read_xlsx(path):
    """Прочитать XLSX-файл и вернуть список словарей."""
    if not os.path.isfile(path):
        print(f"❌ Файл не найден: {path}", file=sys.stderr)
        return None

    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    try:
        with zipfile.ZipFile(path) as archive:
            shared_strings = []
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_root.findall(".//m:si", ns):
                parts = [node.text or "" for node in item.findall(".//m:t", ns)]
                shared_strings.append("".join(parts))

            sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            parsed_rows = {}
            for row in sheet.findall(".//m:sheetData/m:row", ns):
                row_num = int(row.get("r"))
                cells = {}
                for cell in row.findall("m:c", ns):
                    match = re.match(r"([A-Z]+)", cell.get("r"))
                    idx = 0
                    for ch in match.group(1):
                        idx = idx * 26 + (ord(ch) - ord("A") + 1)
                    idx -= 1
                    value_node = cell.find("m:v", ns)
                    if value_node is None:
                        continue
                    raw = value_node.text or ""
                    if cell.get("t") == "s":
                        raw = shared_strings[int(raw)]
                    cells[idx] = str(raw).strip()
                parsed_rows[row_num] = cells
    except (KeyError, OSError, ET.ParseError, ValueError) as exc:
        print(f"❌ Ошибка чтения XLSX {path}: {exc}", file=sys.stderr)
        return None

    if not parsed_rows:
        return []

    header_row_num = min(parsed_rows)
    header_cells = parsed_rows[header_row_num]
    headers = [header_cells.get(i, "").strip() for i in sorted(header_cells)]

    records = []
    for row_num in sorted(parsed_rows):
        if row_num == header_row_num:
            continue
        cells = parsed_rows[row_num]
        record = {}
        empty = True
        for i, header in enumerate(headers):
            if not header:
                continue
            value = cells.get(i, "")
            record[header] = value
            if value:
                empty = False
        if not empty:
            records.append(record)
    return records


def read_spreadsheet(path):
    """Прочитать ODS или XLSX."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".ods":
        return read_ods(path)
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    print(f"❌ Неподдерживаемый формат: {ext}", file=sys.stderr)
    return None


def format_passport_title(record):
    """Заголовок страницы — только название системы."""
    return record_field(record, "название системы") or "Система"


def format_passport_content(record):
    """Содержимое страницы: подпись поля и значение."""
    blocks = []
    for field in PASSPORT_CONTENT_FIELDS:
        label = PASSPORT_FIELD_LABELS.get(field, field)
        value = record_field(record, field) or "—"
        blocks.append(f"**{label}:**\n\n{value}")
    return "\n\n".join(blocks)


def resolve_unique_slug(token, parent_slug, record, used_slugs):
    """Подобрать уникальный slug: при совпадении добавляет суффикс (IP, имя, номер)."""
    base = slugify(record_field(record, "Hostname")) or slugify(record_field(record, "название системы")) or "system"
    ip_part = record_field(record, "IP адрес").replace(".", "-")
    name_part = slugify(record_field(record, "название системы"))

    candidates = [base]
    if ip_part:
        candidates.append(f"{base}-{ip_part}")
    if name_part and name_part != base:
        candidates.append(f"{base}-{name_part}")

    for n in range(2, 100):
        candidates.append(f"{base}-{n}")

    for part in candidates:
        slug = f"{parent_slug.rstrip('/')}/{part}"
        if slug in used_slugs:
            continue
        if token is not None and get_page(token, slug):
            continue
        used_slugs.add(slug)
        return slug

    fallback = f"{parent_slug.rstrip('/')}/{base}-{len(used_slugs) + 1}"
    used_slugs.add(fallback)
    return fallback


def run_create_passports_mode(token, parent_slug, data_path, args):
    """Создать подстраницы-паспорта для каждой системы из файла."""
    print(f"\n📂 Читаю {data_path}...")
    records = read_spreadsheet(data_path)
    if records is None:
        sys.exit(1)
    if not records:
        print("❌ В файле нет данных.")
        sys.exit(1)

    print(f"✅ Найдено систем: {len(records)}")
    print(f"📁 Родительская страница: {parent_slug}\n")

    created = failed = 0
    used_slugs = set()

    for i, record in enumerate(records, 1):
        title = format_passport_title(record)
        content = format_passport_content(record)
        slug = resolve_unique_slug(token, parent_slug, record, used_slugs)

        print(f"[{i}/{len(records)}] {title}")
        print(f"    slug: {slug}")

        if args.dry_run:
            print("    --- содержимое ---")
            for line in content.split("\n\n"):
                preview = line if len(line) <= 60 else line[:60] + "..."
                print(f"    {preview}")
            print()
            continue

        print("    📝 Создаю страницу...")
        result = create_wiki_page(token, slug, title, content)
        if result:
            created += 1
            print(f"    ✅ Создано: https://wiki.yandex.ru/{slug}")
        else:
            failed += 1
        print()

    if args.dry_run:
        print(f"🔍 Dry-run: будет создано {len(records)} страниц")
        return

    print("=" * 60)
    print(f"✅ Создано: {created} | ❌ Ошибок: {failed}")
    print(f"🔗 https://wiki.yandex.ru/{parent_slug}")


def _page_field(page_data, key, default="—"):
    """Поле страницы или из attributes."""
    if key in page_data and page_data[key] is not None:
        return page_data[key]
    attrs = page_data.get("attributes") or {}
    if key in attrs and attrs[key] is not None:
        return attrs[key]
    return default


def print_page_details(page_data, index=None, total=None, content_limit=None):
    """Вывести подробную информацию о странице в консоль."""
    header = "📄 Информация о странице"
    if index is not None and total is not None:
        header = f"📄 Страница {index}/{total}"

    print("\n" + "=" * 60)
    print(header)
    print("=" * 60)
    print(f"  ID:          {page_data.get('id', '—')}")
    print(f"  Заголовок:   {page_data.get('title', '—')}")
    print(f"  Slug:        {page_data.get('slug', '—')}")
    print(f"  URL:         https://wiki.yandex.ru/{page_data.get('slug', '')}")
    print(f"  Тип:         {page_data.get('page_type', '—')}")
    print(f"  Создана:     {_page_field(page_data, 'created_at')}")
    print(f"  Изменена:    {_page_field(page_data, 'modified_at')}")
    print(f"  Комментарии: {_page_field(page_data, 'comments_count', 0)}")
    print(f"  Черновик:    {'Да' if _page_field(page_data, 'is_draft') else 'Нет'}")
    print(f"  Только чтение: {'Да' if _page_field(page_data, 'is_readonly') else 'Нет'}")

    breadcrumbs = page_data.get("breadcrumbs")
    if breadcrumbs:
        path = " → ".join(crumb.get("title") or crumb.get("slug", "?") for crumb in breadcrumbs)
        print(f"  Путь:        {path}")

    redirect = page_data.get("redirect")
    if redirect:
        print(f"  Редирект:    {redirect.get('slug', redirect)}")

    content = page_data.get("content")
    if content:
        print("\n📝 Содержимое:")
        print("-" * 40)
        if content_limit is not None and len(content) > content_limit:
            print(content[:content_limit] + f"\n... (ещё {len(content) - content_limit} символов)")
        else:
            print(content)
    print("=" * 60)


def fetch_pages_details(token, pages, fields="content,attributes,breadcrumbs"):
    """Загрузить подробную информацию для списка страниц."""
    detailed = []
    total = len(pages)
    sorted_pages = sorted(pages, key=lambda p: (p.get("slug", "").count("/"), p.get("slug", "")))

    for i, page in enumerate(sorted_pages, 1):
        slug_page = page.get("slug")
        page_id = page.get("id")
        print(f"📥 Загрузка {i}/{total}: {slug_page}...", file=sys.stderr)

        page_data = None
        if slug_page:
            page_data = get_page(token, slug_page, fields=fields)
        if page_data is None and page_id:
            page_data = get_page_by_id(token, page_id, fields=fields)
        if page_data is None:
            page_data = page

        detailed.append(page_data)

    return detailed


def get_current_user(token):
    """Информация о текущем пользователе."""
    url = "https://api.wiki.yandex.net/v1/users/me"
    try:
        resp = session.get(url, headers=get_headers(token), timeout=30)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        print(f"⚠️  Не удалось получить данные пользователя: {e}", file=sys.stderr)
        return None


def get_descendants(token, slug=None, include_self=False, page_size=100, whole_wiki=False):
    """Получить список потомков с пагинацией."""
    all_pages = []
    cursor = None
    page_num = 1
    base_url = "https://api.wiki.yandex.net/v1/pages/descendants"
    headers = get_headers(token)

    while True:
        # Пустой slug="" — обход всей вики (важно не omit-ить параметр!)
        if whole_wiki:
            params = [("slug", ""), ("page_size", page_size)]
        elif slug:
            params = {"slug": slug, "include_self": include_self, "page_size": page_size}
        else:
            print("❌ Нужен slug или whole_wiki=True.", file=sys.stderr)
            return None

        if cursor:
            if isinstance(params, list):
                params.append(("cursor", cursor))
            else:
                params["cursor"] = cursor

        print(f"📄 Загрузка страницы {page_num}...", file=sys.stderr)

        try:
            resp = session.get(base_url, headers=headers, params=params, timeout=60)
            resp.raise_for_status()
            data = resp.json()

            pages = data.get("results") or []
            all_pages.extend(pages)

            cursor = data.get("next_cursor")
            if not cursor:
                break

            page_num += 1

        except requests.exceptions.RequestException as e:
            print(f"❌ Ошибка при получении потомков: {e}", file=sys.stderr)
            if hasattr(e, 'response') and e.response:
                print(f"   Ответ: {e.response.text}", file=sys.stderr)
            return None

    return all_pages


def search_pages(token, queries=None, page_size=50, filters=None, highlight=True):
    """Поиск страниц через API (fallback для обзора)."""
    if not queries:
        queries = ["*", "a", "и", "страница", "wiki", "ntu"]

    seen_slugs = set()
    all_results = []
    url = "https://api.wiki.yandex.net/v1/search"
    headers = get_headers(token)
    body_filters = {"type": "page"}
    if filters:
        body_filters.update(filters)

    for query in queries:
        cursor = 1
        page_num = 1

        while True:
            body = {
                "query": query,
                "filters": body_filters,
                "limit": min(page_size, 50),
                "cursor": cursor,
                "order_by": "modified_date",
                "highlight": highlight,
            }

            print(f"🔍 Поиск «{query}», страница {page_num}...", file=sys.stderr)

            try:
                resp = session.post(url, headers=headers, json=body, timeout=30)
                resp.raise_for_status()
                data = resp.json()

                results = data.get("results") or []
                if not results:
                    break

                for item in results:
                    slug = item.get("slug")
                    if not slug or slug in seen_slugs:
                        continue
                    seen_slugs.add(slug)
                    all_results.append({
                        "id": item.get("id"),
                        "slug": slug,
                        "title": item.get("title"),
                        "modified_at": item.get("modified_at"),
                    })

                if not highlight:
                    break

                next_cursor = data.get("next_cursor")
                if not next_cursor:
                    break
                try:
                    cursor = int(next_cursor)
                except (TypeError, ValueError):
                    break

                page_num += 1

            except requests.exceptions.RequestException as e:
                print(f"⚠️  Поиск «{query}» не удался: {e}", file=sys.stderr)
                break

    return all_results


def get_all_accessible_pages(token, page_size=100):
    """Все страницы вики, доступные пользователю."""
    pages = get_descendants(token, whole_wiki=True, page_size=page_size)
    if pages:
        return pages

    print("ℹ️  Обход по корню не дал результатов, пробую поиск...", file=sys.stderr)
    pages = search_pages(token, page_size=page_size)
    return pages


def ask_slug(prompt="Введите URL или slug страницы"):
    """Запросить slug/URL у пользователя."""
    print(f"\n📌 {prompt}.")
    print("   Пример URL: https://wiki.yandex.ru/ntu/900f18c0fde1/pasporta/")
    print("   Пример slug: ntu/900f18c0fde1/pasporta")
    user_input = input("> ").strip()
    if not user_input:
        print("❌ Ничего не введено.")
        return None
    return normalize_slug(user_input)


def print_pages_overview(pages, max_items_per_section=30):
    """Обзор доступных страниц, сгруппированный по разделам."""
    if not pages:
        print("📭 Страницы не найдены.")
        return

    sections = {}
    for page in pages:
        slug = page.get("slug") or ""
        section = slug.split("/")[0] if slug else "?"
        sections.setdefault(section, []).append(page)

    print(f"\n📚 Доступно страниц: {len(pages)}")
    print(f"📁 Разделов верхнего уровня: {len(sections)}\n")
    print("=" * 60)

    for section in sorted(sections):
        items = sorted(sections[section], key=lambda p: p.get("slug", ""))
        print(f"\n📁 {section}/ — {len(items)} стр.")
        print("-" * 60)
        shown = items[:max_items_per_section]
        for page in shown:
            slug = page.get("slug", "—")
            title = page.get("title")
            page_id = page.get("id", "—")
            depth = slug.count("/") if slug != "—" else 0
            indent = "  " * depth
            if title:
                print(f"{indent}• {title}")
                print(f"{indent}  {slug} (ID: {page_id})")
            else:
                print(f"{indent}• {slug} (ID: {page_id})")
        if len(items) > max_items_per_section:
            print(f"  ... и ещё {len(items) - max_items_per_section} страниц в разделе")

    print("\n" + "=" * 60)
    print(f"✅ Всего: {len(pages)} страниц в {len(sections)} разделах")
    print("💡 Чтобы посмотреть подстраницы раздела: выберите пункт 1 и введите slug раздела")


def print_pages_list_only(pages):
    """Краткий список slug/ID."""
    print(f"\n📚 Найдено страниц: {len(pages)}\n")
    print("-" * 60)
    for page in sorted(pages, key=lambda p: p.get("slug", "")):
        page_id = page.get("id", "—")
        slug_page = page.get("slug", "—")
        title = page.get("title")
        if title:
            print(f"• {title}")
            print(f"  {slug_page} (ID: {page_id})")
        else:
            print(f"- {slug_page} (ID: {page_id})")
    print(f"\nВсего: {len(pages)}")


def output_pages(token, pages, args):
    """Вывести страницы в выбранном формате."""
    if not pages:
        print("📭 Страницы не найдены.")
        return

    if args.raw:
        detailed = fetch_pages_details(token, pages)
        print(json.dumps(detailed, indent=2, ensure_ascii=False))
    elif args.list_only:
        print_pages_list_only(pages)
    else:
        print(f"\n📂 Загружаю подробную информацию о {len(pages)} страницах...")
        detailed = fetch_pages_details(token, pages)
        print(f"\n📚 Найдено страниц: {len(detailed)}\n")
        for i, page_data in enumerate(detailed, 1):
            print_page_details(page_data, i, len(detailed), args.content_limit)
        print(f"\n✅ Всего обработано: {len(detailed)}")


def run_subpages_mode(token, slug, args):
    """Подстраницы указанной ссылки."""
    print(f"\n📂 Получение подстраниц для '{slug}' (include_self={args.include_self})...")
    pages = get_descendants(token, slug, args.include_self, args.page_size)
    if pages is None:
        print("❌ Ошибка при получении списка страниц.")
        sys.exit(1)
    output_pages(token, pages, args)


def run_single_page_mode(token, slug, args):
    """Подробная информация об одной странице."""
    print(f"\n📄 Получение информации о странице '{slug}'...")
    page_data = get_page(token, slug, fields="content,attributes,breadcrumbs")
    if page_data is None:
        print("❌ Не удалось получить страницу. Проверьте правильность slug.")
        sys.exit(1)
    if args.raw:
        print(json.dumps(page_data, indent=2, ensure_ascii=False))
    else:
        print_page_details(page_data, content_limit=args.content_limit)


def run_all_pages_mode(token, args):
    """Обзор всех доступных страниц."""
    user = get_current_user(token)
    if user:
        print(f"\n👤 Пользователь: {user.get('username', '—')}")

    print("\n🌐 Загружаю все доступные страницы...")
    pages = get_all_accessible_pages(token, args.page_size)
    if pages is None:
        print("❌ Ошибка при получении списка страниц.")
        sys.exit(1)

    if not pages:
        print("📭 Страницы не найдены. Возможно, нет доступа или вики пуста.")
        return

    if args.raw:
        print(json.dumps(pages, indent=2, ensure_ascii=False))
    elif args.list_only:
        print_pages_list_only(pages)
    else:
        print_pages_overview(pages)
        if len(pages) <= 20:
            answer = input("\n❓ Загрузить подробную информацию по всем? [y/N]: ").strip().lower()
            if answer in ("y", "yes", "д", "да"):
                output_pages(token, pages, args)


def show_interactive_menu(args):
    """Интерактивное меню выбора режима."""
    print("\n" + "=" * 60)
    print("  Яндекс Вики — выберите действие")
    print("=" * 60)
    print("  1 — Создать паспорта систем из файла")
    print("  2 — Подстраницы по ссылке (slug/URL)")
    print("  3 — Все доступные страницы (обзор)")
    print("  4 — Одна страница (подробно)")
    print("  0 — Выход")
    print("=" * 60)

    choice = input("\n> ").strip()
    if choice == "0":
        print("👋 До свидания.")
        sys.exit(0)
    if choice == "1":
        return "create", DEFAULT_PARENT_SLUG
    if choice == "2":
        return "subpages", ask_slug()
    if choice == "3":
        return "all", None
    if choice == "4":
        return "single", ask_slug()
    print("❌ Неверный выбор.")
    return None, None


def get_token():
    """Получить OAuth-токен."""
    token = os.environ.get("WIKI_TOKEN")
    if token:
        print("✅ Использую токен из переменной окружения WIKI_TOKEN")
    else:
        print("🔐 Начинаю процесс авторизации...")
        token = get_oauth_token(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI)
    print(f"✅ Использую организацию с ID: {ORG_ID}")
    return token


def main():
    parser = argparse.ArgumentParser(
        description="Создание паспортов систем в Яндекс Вики из ODS/XLSX"
    )
    parser.add_argument("--create", "-c", action="store_true",
                        help="Создать подстраницы-паспорта (режим по умолчанию)")
    parser.add_argument("--file", "-f", type=str, default=DEFAULT_DATA_FILE,
                        help=f"Файл с данными ODS/XLSX (по умолчанию: {DEFAULT_DATA_FILE})")
    parser.add_argument("--parent-slug", type=str, default=DEFAULT_PARENT_SLUG,
                        help=f"Родительская страница (по умолчанию: {DEFAULT_PARENT_SLUG})")
    parser.add_argument("--dry-run", action="store_true",
                        help="Показать что будет создано, без записи в вики")
    parser.add_argument("--url", type=str, help="Полный URL родительской страницы")
    parser.add_argument("--slug", type=str, help="Slug родительской страницы")
    parser.add_argument("--info", action="store_true",
                        help="Показать информацию о странице")
    parser.add_argument("--all-pages", "-a", action="store_true",
                        help="Показать все доступные страницы (обзор)")
    parser.add_argument("--menu", "-m", action="store_true",
                        help="Интерактивное меню")
    parser.add_argument("--descendants", "-d", action="store_true",
                        help="Показать список дочерних страниц")
    parser.add_argument("--include-self", "-s", action="store_true", default=True,
                        help="Включить саму страницу в список потомков")
    parser.add_argument("--no-include-self", action="store_false", dest="include_self",
                        help="Не включать саму страницу в список потомков")
    parser.add_argument("--page-size", type=int, default=100,
                        help="Количество страниц на одну загрузку")
    parser.add_argument("--list-only", action="store_true",
                        help="Только список slug/ID без подробной информации")
    parser.add_argument("--content-limit", type=int, default=None,
                        help="Ограничить длину содержимого при просмотре")
    parser.add_argument("--raw", action="store_true",
                        help="Вывести сырой JSON-ответ")
    args = parser.parse_args()

    read_modes = args.info or args.all_pages or args.descendants
    use_menu = args.menu

    if use_menu:
        token = get_token()
        while True:
            mode, slug = show_interactive_menu(args)
            if mode is None:
                continue
            if mode == "create":
                run_create_passports_mode(token, slug or DEFAULT_PARENT_SLUG, args.file, args)
            elif mode == "subpages":
                if not slug:
                    continue
                print(f"🔍 Использую slug: {slug}")
                run_subpages_mode(token, slug, args)
            elif mode == "all":
                run_all_pages_mode(token, args)
            elif mode == "single":
                if not slug:
                    continue
                print(f"🔍 Использую slug: {slug}")
                run_single_page_mode(token, slug, args)
            again = input("\n❓ Ещё действие? [Y/n]: ").strip().lower()
            if again in ("n", "no", "н", "нет"):
                break
        return

    if read_modes:
        slug = normalize_slug(args.url) if args.url else normalize_slug(args.slug or DEFAULT_PARENT_SLUG)
        if args.url:
            print(f"🔍 Извлечён slug из URL: {slug}")
        else:
            print(f"🔍 Использую slug: {slug}")
        token = get_token()
        if args.all_pages:
            run_all_pages_mode(token, args)
        elif args.info:
            run_single_page_mode(token, slug, args)
        else:
            run_subpages_mode(token, slug, args)
        return

    parent_slug = normalize_slug(args.url) if args.url else normalize_slug(args.parent_slug)
    print(f"📁 Родитель: {parent_slug}")
    print(f"📊 Файл: {args.file}")

    if args.dry_run:
        run_create_passports_mode(None, parent_slug, args.file, args)
        return

    token = get_token()
    run_create_passports_mode(token, parent_slug, args.file, args)


if __name__ == "__main__":
    main()