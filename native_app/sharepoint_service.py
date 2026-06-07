from __future__ import annotations

from dataclasses import dataclass
from html import unescape
import re
import tempfile
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qs, quote, urljoin, urlparse, unquote

import requests

from src.revenue_loader import FILE_PATTERN


HREF_PATTERN = re.compile(r'href=["\']([^"\']+)["\']', re.IGNORECASE)
FILE_NAME_PATTERN = re.compile(r'(?P<name>[A-Za-z0-9][A-Za-z0-9 _.-]*\.(?:xlsx|xlsm|xls))', re.IGNORECASE)
DEBUG_LOG_PATH = Path(tempfile.gettempdir()) / "revenue_sharepoint_debug.log"
AUTH_COOKIE_NAMES = {"fedauth", "rtfa", "spoidcrl"}


def _append_debug_log(message: str) -> None:
    try:
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(message + "\n")
    except Exception:
        pass


@dataclass(frozen=True)
class SharePointCookieRecord:
    name: str
    value: str
    domain: str = ""
    path: str = "/"
    secure: bool = False


@dataclass(frozen=True)
class SharePointFileRef:
    name: str
    url: str
    kind: str


@dataclass(frozen=True)
class SharePointDownloadResult:
    data_folder: Path
    template_file: Path
    discovered_files: tuple[SharePointFileRef, ...]


def build_requests_session(cookies: Iterable[SharePointCookieRecord]) -> requests.Session:
    session = requests.Session()
    cookie_list = [cookie for cookie in cookies if cookie.name and cookie.value]
    auth_cookies = [cookie for cookie in cookie_list if cookie.name.lower() in AUTH_COOKIE_NAMES]
    effective_cookies = auth_cookies or cookie_list

    header_parts: list[str] = []
    for cookie in effective_cookies:
        header_parts.append(f"{cookie.name}={cookie.value}")
        session.cookies.set(
            cookie.name,
            cookie.value,
            domain=(cookie.domain.lstrip(".") or None),
            path=cookie.path or "/",
            secure=cookie.secure,
        )
    if header_parts:
        # Reference apps send the raw Cookie header directly instead of relying
        # on cookie-jar domain/path reconstruction.
        session.headers.update({"Cookie": "; ".join(dict.fromkeys(header_parts))})
    _append_debug_log(
        f"sharepoint_session_built | auth_cookie_count={len(auth_cookies)} | effective_cookie_count={len(effective_cookies)}"
    )
    return session


def _link_name(link_url: str, original_href: str) -> str:
    parsed = urlparse(link_url)
    file_name = unquote(Path(parsed.path).name)
    if file_name:
        return file_name

    query = parse_qs(parsed.query)
    for key in ("SourceUrl", "sourceUrl", "src", "file"):
        values = query.get(key)
        if values:
            candidate = unquote(Path(urlparse(values[0]).path).name)
            if candidate:
                return candidate

    fallback = unquote(Path(urlparse(original_href).path).name)
    return fallback or "sharepoint-file.xlsx"


def _normalize_folder_path(folder_url: str) -> tuple[str, str, str]:
    parsed = urlparse(folder_url)
    raw_path = unquote(parsed.path)
    match = re.match(r"^/:[^/]+/r(?P<relative>/.*)$", raw_path, re.IGNORECASE)
    server_relative_path = match.group("relative") if match else raw_path
    server_relative_path = server_relative_path.rstrip("/") or "/"
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    return base_url, server_relative_path, parsed.netloc


def _folder_url_from_server_relative(base_url: str, server_relative_path: str) -> str:
    normalized = server_relative_path.rstrip("/") or "/"
    return urljoin(base_url, quote(normalized, safe="/%"))


def _normalize_extracted_url(raw: str) -> str:
    return (
        unescape(raw)
        .replace("\\u002F", "/")
        .replace("\\u0026", "&")
        .replace("\\u003A", ":")
        .replace("\\u003F", "?")
        .replace("\\u003D", "=")
        .replace("\\u0025", "%")
        .replace("\\/", "/")
        .strip("\"' ")
    )


def _extract_download_urls(base_url: str, html: str) -> list[str]:
    patterns = [
        r'"FileGetUrl"\s*:\s*"(?P<url>[^"]+)"',
        r'"FileUrlNoAuth"\s*:\s*"(?P<url>[^"]+)"',
        r'"downloadUrl"\s*:\s*"(?P<url>[^"]+)"',
        r'https?:\\/\\/[^"\'\s>]*download\.aspx[^"\'\s>]*',
        r'https?://[^"\'\s>]*download\.aspx[^"\'\s>]*',
        r'/_layouts/15/download\.aspx[^"\'\s>]*',
        r'https?:\\/\\/[^"\'\s>]*\.xls[xm]?[^"\'\s>]*',
        r'https?://[^"\'\s>]*\.xls[xm]?[^"\'\s>]*',
        r'/personal/[^"\'\s>]*\.xls[xm]?[^"\'\s>]*',
    ]

    discovered: list[str] = []
    seen: set[str] = set()
    for pattern in patterns:
        for match in re.finditer(pattern, html, re.IGNORECASE):
            raw = match.groupdict().get("url") or match.group(0)
            normalized = _normalize_extracted_url(raw).rstrip(",;")
            if not normalized:
                continue
            if (
                "download.aspx" not in normalized.lower()
                and not normalized.lower().endswith((".xlsx", ".xlsm", ".xls"))
                and ".xlsx?" not in normalized.lower()
                and ".xlsm?" not in normalized.lower()
            ):
                continue

            absolute = urljoin(base_url, normalized)
            if absolute.lower() in seen:
                continue
            seen.add(absolute.lower())
            discovered.append(absolute)
    return discovered


def _recover_file_refs_from_html(folder_url: str, html: str) -> list[SharePointFileRef]:
    base_url, server_relative_path, _ = _normalize_folder_path(folder_url)
    text = unescape(html)
    discovered: list[SharePointFileRef] = []
    seen: set[str] = set()
    for match in FILE_NAME_PATTERN.finditer(text):
        name = match.group("name").strip()
        if len(name) > 160:
            continue
        lower_name = name.lower()
        if lower_name in seen:
            continue
        seen.add(lower_name)

        relative_file_path = f"{server_relative_path.rstrip('/')}/{name}"
        absolute_url = urljoin(base_url, quote(relative_file_path, safe='/%'))
        kind = "monthly" if FILE_PATTERN.match(name) else "template"
        discovered.append(SharePointFileRef(name=name, url=absolute_url, kind=kind))

    discovered.sort(key=lambda item: (0 if item.kind == "template" else 1, item.name.lower()))
    return discovered


def _discover_sharepoint_files_via_rest(folder_url: str, session: requests.Session, timeout: int = 60) -> list[SharePointFileRef]:
    base_url, server_relative_path, host = _normalize_folder_path(folder_url)
    escaped_path = server_relative_path.replace("'", "''")
    candidate_urls = [
        f"{base_url}/_api/web/GetFolderByServerRelativeUrl('{escaped_path}')/Files?$select=Name,ServerRelativeUrl",
        f"{base_url}/_api/web/GetFolderByServerRelativePath(decodedurl='{escaped_path}')/Files?$select=Name,ServerRelativeUrl",
    ]

    results: list[dict] = []
    last_error: Exception | None = None
    for api_url in candidate_urls:
        try:
            response = session.get(
                api_url,
                timeout=timeout,
                headers={"Accept": "application/json;odata=verbose"},
            )
            _append_debug_log(
                f"sharepoint_rest_list | host={host} | folder={server_relative_path} | status={response.status_code} | url={response.url}"
            )
            response.raise_for_status()

            payload = response.json()
            results = payload.get("d", {}).get("results", [])
            if results:
                break
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            _append_debug_log(f"sharepoint_rest_candidate_failed | url={api_url} | error={exc}")

    if not results and last_error is not None:
        raise last_error

    discovered: list[SharePointFileRef] = []
    for item in results:
        name = str(item.get("Name") or "").strip()
        relative_url = str(item.get("ServerRelativeUrl") or "").strip()
        if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        if not relative_url:
            continue

        absolute_url = urljoin(base_url, quote(relative_url, safe="/%"))
        kind = "monthly" if FILE_PATTERN.match(name) else "template"
        discovered.append(SharePointFileRef(name=name, url=absolute_url, kind=kind))

    _append_debug_log(
        f"sharepoint_rest_results | count={len(discovered)} | names={', '.join(item.name for item in discovered[:20]) or 'none'}"
    )
    discovered.sort(key=lambda item: (0 if item.kind == "template" else 1, item.name.lower()))
    return discovered


def _discover_sharepoint_files_via_html(folder_url: str, session: requests.Session, timeout: int = 60) -> list[SharePointFileRef]:
    response = session.get(folder_url, timeout=timeout)
    _append_debug_log(
        f"sharepoint_html_list | status={response.status_code} | url={response.url} | content_type={response.headers.get('Content-Type', '')}"
    )
    response.raise_for_status()

    discovered: list[SharePointFileRef] = []
    seen_urls: set[str] = set()
    for href in HREF_PATTERN.findall(response.text):
        absolute_url = urljoin(folder_url, unescape(href))
        lower_url = absolute_url.lower()
        if ".xlsx" not in lower_url and "download.aspx" not in lower_url:
            continue

        name = _link_name(absolute_url, href)
        if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        if absolute_url in seen_urls:
            continue

        seen_urls.add(absolute_url)
        kind = "monthly" if FILE_PATTERN.match(name) else "template"
        discovered.append(SharePointFileRef(name=name, url=absolute_url, kind=kind))

    discovered.sort(key=lambda item: (0 if item.kind == "template" else 1, item.name.lower()))
    if discovered:
        _append_debug_log(
            f"sharepoint_html_results | count={len(discovered)} | names={', '.join(item.name for item in discovered[:20]) or 'none'}"
        )
        return discovered

    download_urls = _extract_download_urls(response.url, response.text)
    for absolute_url in download_urls:
        name = _link_name(absolute_url, absolute_url)
        if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        kind = "monthly" if FILE_PATTERN.match(name) else "template"
        discovered.append(SharePointFileRef(name=name, url=absolute_url, kind=kind))

    if not discovered:
        discovered = _recover_file_refs_from_html(folder_url, response.text)

    discovered.sort(key=lambda item: (0 if item.kind == "template" else 1, item.name.lower()))
    _append_debug_log(
        f"sharepoint_html_results | count={len(discovered)} | names={', '.join(item.name for item in discovered[:20]) or 'none'}"
    )
    return discovered


def discover_sharepoint_files(folder_url: str, session: requests.Session, timeout: int = 60) -> list[SharePointFileRef]:
    try:
        discovered = _discover_sharepoint_files_via_rest(folder_url, session, timeout=timeout)
        if discovered:
            return discovered
        _append_debug_log("sharepoint_rest_empty | falling_back_to_html=true")
    except Exception as exc:  # noqa: BLE001
        _append_debug_log(f"sharepoint_rest_failed | error={exc}")

    return _discover_sharepoint_files_via_html(folder_url, session, timeout=timeout)


def _build_download_url(file_url: str) -> str:
    parsed = urlparse(file_url)
    if parsed.path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return file_url

    source_url = quote(file_url, safe="")
    return f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?SourceUrl={source_url}"


def _build_download_candidates(file_url: str) -> list[str]:
    parsed = urlparse(file_url)
    candidates = [file_url]

    if "download=1" not in parsed.query.lower():
        separator = "&" if parsed.query else "?"
        candidates.append(f"{file_url}{separator}download=1")

    source_url = quote(file_url, safe="")
    candidates.append(f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?SourceUrl={source_url}")
    candidates.append(f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?sourceurl={source_url}")

    deduped: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        if candidate.lower() in seen:
            continue
        seen.add(candidate.lower())
        deduped.append(candidate)
    return deduped


def _is_html_response(content: bytes, media_type: str) -> bool:
    if "html" in media_type.lower():
        return True
    prefix = content[:256].decode("utf-8", errors="ignore").lstrip()
    return prefix.lower().startswith("<!doctype html") or prefix.lower().startswith("<html")


def _is_spreadsheet_response(content: bytes, media_type: str, url: str) -> bool:
    if _is_html_response(content, media_type):
        return False
    if len(content) >= 2 and content[0] == 0x50 and content[1] == 0x4B:
        return True
    lowered = media_type.lower()
    return (
        "spreadsheet" in lowered
        or "excel" in lowered
        or "octet-stream" in lowered
        or urlparse(url).path.lower().endswith((".xlsx", ".xlsm", ".xls"))
    )


def download_sharepoint_file(session: requests.Session, file_url: str, output_path: Path, timeout: int = 120) -> None:
    candidate_urls = _build_download_candidates(file_url)
    last_error: Exception | None = None

    for candidate_url in list(candidate_urls):
        try:
            response = session.get(candidate_url, timeout=timeout)
            media_type = response.headers.get("Content-Type", "")
            _append_debug_log(
                f"sharepoint_download | status={response.status_code} | url={response.url} | media_type={media_type} | output={output_path.name}"
            )
            response.raise_for_status()
            content = response.content

            if _is_spreadsheet_response(content, media_type, str(response.url)):
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(content)
                return

            if _is_html_response(content, media_type):
                html = content.decode("utf-8", errors="ignore")
                for extra_url in _extract_download_urls(str(response.url), html):
                    if extra_url.lower() not in {url.lower() for url in candidate_urls}:
                        candidate_urls.append(extra_url)
                last_error = ValueError(f"SharePoint returned an HTML preview page for {candidate_url}")
                continue

            last_error = ValueError(f"Downloaded content from {candidate_url} was not recognized as an Excel workbook.")
        except Exception as exc:  # noqa: BLE001
            last_error = exc

    if last_error is not None:
        raise last_error
    raise ValueError(f"Unable to download SharePoint file '{file_url}'.")


def download_sharepoint_dataset(
    folder_url: str,
    cookies: Iterable[SharePointCookieRecord],
    template_name: str = "Revenue Output Dashboard Sample.xlsx",
) -> SharePointDownloadResult:
    session = build_requests_session(cookies)
    discovered = discover_sharepoint_files(folder_url, session)
    if not discovered:
        raise ValueError("No Excel files were discovered at the provided SharePoint folder URL.")

    monthly_files = [item for item in discovered if FILE_PATTERN.match(item.name)]
    if not monthly_files:
        raise ValueError("No monthly revenue files matched the expected filename pattern.")

    template_candidates = [
        item
        for item in discovered
        if item.name.lower() == template_name.lower() or "template" in item.name.lower()
    ]
    template_item = template_candidates[0] if template_candidates else None
    if template_item is None:
        base_url, server_relative_path, _ = _normalize_folder_path(folder_url)
        parent_relative_path = str(Path(server_relative_path).parent).replace("\\", "/")
        if parent_relative_path and parent_relative_path != "." and parent_relative_path != server_relative_path:
            parent_folder_url = _folder_url_from_server_relative(base_url, parent_relative_path)
            _append_debug_log(
                f"sharepoint_template_parent_lookup | parent_folder={parent_relative_path} | url={parent_folder_url}"
            )
            parent_discovered = discover_sharepoint_files(parent_folder_url, session)
            parent_template_candidates = [
                item
                for item in parent_discovered
                if item.name.lower() == template_name.lower() or "template" in item.name.lower()
            ]
            if parent_template_candidates:
                template_item = parent_template_candidates[0]
                discovered = discovered + [
                    item for item in parent_discovered if item.name.lower() == template_item.name.lower()
                ]
    if template_item is None:
        raise ValueError(f"Template file '{template_name}' was not found in the SharePoint folder.")

    temp_root = Path(tempfile.mkdtemp(prefix="revenue-sharepoint-"))
    monthly_dir = temp_root / "monthly"
    monthly_dir.mkdir(parents=True, exist_ok=True)

    for item in monthly_files:
        download_sharepoint_file(session, item.url, monthly_dir / item.name)

    template_path = temp_root / template_item.name
    download_sharepoint_file(session, template_item.url, template_path)

    return SharePointDownloadResult(
        data_folder=monthly_dir,
        template_file=template_path,
        discovered_files=tuple(discovered),
    )
