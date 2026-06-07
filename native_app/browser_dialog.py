from __future__ import annotations

from dataclasses import dataclass
import tempfile
from pathlib import Path
from urllib.parse import urlparse

from PySide6.QtCore import QTimer, QUrl
from PySide6.QtWebEngineCore import QWebEngineCookieStore
from PySide6.QtWidgets import (
    QDialog, QHBoxLayout, QLabel,
    QPushButton, QTextBrowser, QVBoxLayout,
)
from PySide6.QtWebEngineWidgets import QWebEngineView

_SHAREPOINT_AUTH_COOKIES = {"fedauth", "rtfa", "spoidcrl"}
DEBUG_LOG_PATH = Path(tempfile.gettempdir()) / "revenue_sharepoint_debug.log"


def _append_debug_log(message: str) -> None:
    try:
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(message + "\n")
    except Exception:
        pass


def _safe_login_url(sharing_url: str) -> str:
    """
    Replicates AllocationReview's GetSafeSharePointLoginUrl logic.
    Strips the sharing link down to the SharePoint host + /personal/<user>/ root
    so the embedded browser navigates to the right domain for cookie auth.
    """
    try:
        parsed = urlparse(sharing_url)
        segments = [s for s in parsed.path.split("/") if s]
        for index, segment in enumerate(segments):
            lowered = segment.lower()
            if lowered in {"sites", "teams", "personal"} and len(segments) > index + 1:
                return f"{parsed.scheme}://{parsed.netloc}/{segments[index]}/{segments[index + 1]}/"
        return f"{parsed.scheme}://{parsed.netloc}/"
    except Exception:
        return sharing_url


@dataclass(frozen=True)
class CookieSnapshot:
    name: str
    value: str
    domain: str
    path: str
    secure: bool


class SharePointLoginDialog(QDialog):
    """
    Mirrors the pattern used in all three reference apps (AllocationReview,
    NewAllocationTemplate, TimesheetDefaulterTracker):

    1. Navigate to safe site-root URL so Microsoft sets auth cookies on the
       correct domain (GetSafeSharePointLoginUrl equivalent).
    2. On every NavigationCompleted, call GetCookiesAsync(currentUrl) — scoped
       to the current page URL — and check for FedAuth / rtFa / SPOIDCRL.
       Qt equivalent: inject JS document.cookie read + use cookieStore signals.
    3. Auto-close the moment those cookies are detected.
    4. Manual "Check Session" button as fallback (CaptureSessionButton equivalent).
    5. "Proceed Anyway" if cookies exist but auth names not yet found.
    """

    def __init__(self, sharing_url: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("SharePoint Sign-In")
        self.resize(1280, 860)

        self._sharing_url = sharing_url
        self._login_url = _safe_login_url(sharing_url)
        self._cookies: dict[tuple[str, str, str], CookieSnapshot] = {}
        self._captured_cookies: list[CookieSnapshot] = []
        _append_debug_log(f"dialog_initialized | sharing_url={sharing_url} | login_url={self._login_url}")

        self._browser = QWebEngineView(self)

        # Wire cookie store BEFORE first navigation
        profile = self._browser.page().profile()
        self._cookie_store: QWebEngineCookieStore = profile.cookieStore()
        self._cookie_store.cookieAdded.connect(self._on_cookie_added)
        self._cookie_store.loadAllCookies()

        # Navigate to safe root — same as GetSafeSharePointLoginUrl
        self._browser.setUrl(QUrl(self._login_url))

        # After every page load check cookies (NavigationCompleted equivalent)
        self._browser.loadFinished.connect(self._on_navigation_completed)

        self._status = QLabel(
            f"Sign in at {self._login_url}",
            self,
        )
        self._status.setWordWrap(True)
        self._cookie_info = QLabel("Cookies captured: 0", self)

        self._help = QTextBrowser(self)
        self._help.setMaximumHeight(46)
        self._help.setPlainText(
            "Sign in with your Microsoft account. The dialog closes automatically "
            "when SharePoint auth cookies are detected. If it stays open after "
            "you see the folder page, click 'Check Session'."
        )

        self._check_button = QPushButton("Check Session", self)
        self._proceed_button = QPushButton("Proceed Anyway", self)
        self._proceed_button.setEnabled(False)
        self._cancel_button = QPushButton("Cancel", self)

        btn_row = QHBoxLayout()
        btn_row.addWidget(self._check_button)
        btn_row.addWidget(self._proceed_button)
        btn_row.addStretch(1)
        btn_row.addWidget(self._cancel_button)

        layout = QVBoxLayout(self)
        layout.addWidget(self._status)
        layout.addWidget(self._cookie_info)
        layout.addWidget(self._help)
        layout.addWidget(self._browser, 1)
        layout.addLayout(btn_row)

        self._check_button.clicked.connect(self._manual_check)
        self._proceed_button.clicked.connect(self._proceed_anyway)
        self._cancel_button.clicked.connect(self.reject)

    # ------------------------------------------------------------------ #
    #  Cookie store signal — fires for every added/updated cookie         #
    # ------------------------------------------------------------------ #

    def _on_cookie_added(self, cookie) -> None:
        name = bytes(cookie.name()).decode("utf-8", errors="ignore")
        value = bytes(cookie.value()).decode("utf-8", errors="ignore")
        snapshot = CookieSnapshot(
            name=name, value=value,
            domain=cookie.domain(), path=cookie.path(),
            secure=bool(cookie.isSecure()),
        )
        self._cookies[(snapshot.domain, snapshot.path, snapshot.name)] = snapshot
        count = len(self._cookies)
        self._cookie_info.setText(
            f"Cookies captured: {count}  |  domains: "
            + ", ".join(sorted({c.domain.lstrip(".") for c in self._cookies.values() if c.domain}))
        )
        if count > 0:
            self._proceed_button.setEnabled(True)
        # Eagerly check on each new cookie — catches the moment FedAuth arrives
        if name.lower() in _SHAREPOINT_AUTH_COOKIES:
            _append_debug_log(
                f"auth_cookie_added | name={name} | domain={snapshot.domain} | path={snapshot.path} | secure={snapshot.secure}"
            )
            QTimer.singleShot(200, self._try_capture)

    # ------------------------------------------------------------------ #
    #  NavigationCompleted equivalent                                      #
    # ------------------------------------------------------------------ #

    def _on_navigation_completed(self, ok: bool) -> None:
        _append_debug_log(
            f"navigation_completed | ok={ok} | url={self._browser.url().toString()} | cookie_count={len(self._cookies)}"
        )
        if not ok:
            return
        # Reload all cookies from the profile store for current domain
        self._cookie_store.loadAllCookies()
        # Also read document.cookie via JS — mirrors GetCookiesAsync(currentUrl)
        self._browser.page().runJavaScript(
            "document.cookie",
            self._on_js_cookies,
        )
        # Settle delay then try
        QTimer.singleShot(800, self._try_capture)

    def _current_host(self) -> str:
        return urlparse(self._browser.url().toString()).netloc.lower().lstrip(".")

    def _current_page_cookies(self) -> list[CookieSnapshot]:
        current_host = self._current_host()
        if not current_host:
            return list(self._cookies.values())

        scoped: list[CookieSnapshot] = []
        for cookie in self._cookies.values():
            domain = cookie.domain.lower().lstrip(".")
            if not domain:
                continue
            if current_host == domain or current_host.endswith(f".{domain}") or domain.endswith(current_host):
                scoped.append(cookie)
        return scoped

    def _on_js_cookies(self, js_result: object) -> None:
        """Parse document.cookie JS result and add any missing cookies to store."""
        if not isinstance(js_result, str) or not js_result.strip():
            _append_debug_log(
                f"js_cookies_empty | url={self._browser.url().toString()} | cookie_count={len(self._cookies)}"
            )
            return
        current_url = self._browser.url().toString()
        try:
            domain = urlparse(current_url).netloc
        except Exception:
            domain = ""
        for part in js_result.split(";"):
            part = part.strip()
            if "=" in part:
                name, _, value = part.partition("=")
                name = name.strip()
                value = value.strip()
                key = (domain, "/", name)
                if key not in self._cookies:
                    self._cookies[key] = CookieSnapshot(
                        name=name, value=value, domain=domain, path="/", secure=False
                    )
                    if name.lower() in _SHAREPOINT_AUTH_COOKIES:
                        QTimer.singleShot(200, self._try_capture)
        _append_debug_log(
            f"js_cookies_loaded | url={current_url} | js_cookie_names={','.join(sorted(part.partition('=')[0].strip() for part in js_result.split(';') if '=' in part)) or 'none'}"
        )

    # ------------------------------------------------------------------ #
    #  Core capture logic — GetCookiesAsync(currentUrl) equivalent        #
    # ------------------------------------------------------------------ #

    def _try_capture(self, force_status: bool = False) -> bool:
        """
        Mirrors NewAllocationTemplate CaptureSessionButton + AllocationReview
        TryCaptureSessionAsync: look for FedAuth/rtFa/SPOIDCRL and accept.
        """
        all_cookies = list(self._cookies.values())
        if not all_cookies:
            _append_debug_log(
                f"try_capture | force_status={force_status} | url={self._browser.url().toString()} | all_cookie_count=0"
            )
            return False

        scoped_cookies = self._current_page_cookies()
        all_auth_cookies = [c for c in all_cookies if c.name.lower() in _SHAREPOINT_AUTH_COOKIES]
        auth_cookies = [c for c in scoped_cookies if c.name.lower() in _SHAREPOINT_AUTH_COOKIES]
        _append_debug_log(
            " | ".join(
                [
                    "try_capture",
                    f"force_status={force_status}",
                    f"url={self._browser.url().toString()}",
                    f"current_host={self._current_host()}",
                    f"all_cookie_count={len(all_cookies)}",
                    f"scoped_cookie_count={len(scoped_cookies)}",
                    f"auth_scoped={','.join(sorted({c.name for c in auth_cookies})) or 'none'}",
                    f"auth_all={','.join(sorted({c.name for c in all_auth_cookies})) or 'none'}",
                ]
            )
        )
        if auth_cookies:
            self._captured_cookies = all_cookies
            names_found = ", ".join(sorted({c.name for c in auth_cookies}))
            self._status.setText(
                f"SharePoint session captured ({names_found}). Continuing…"
            )
            self.accept()
            return True

        if all_auth_cookies:
            self._captured_cookies = all_cookies
            names_found = ", ".join(sorted({c.name for c in all_auth_cookies}))
            self._status.setText(
                f"SharePoint session captured ({names_found}). Continuing…"
            )
            self.accept()
            return True

        # Match New Allocation Template's capture button behavior:
        # if the current loaded page has any cookies at all, allow manual capture.
        current_url = self._browser.url().toString().lower()
        on_sharepoint_page = "sharepoint.com" in current_url
        if force_status and (scoped_cookies or (on_sharepoint_page and all_cookies)):
            self._captured_cookies = all_cookies
            self._status.setText(
                f"Captured {len(scoped_cookies) or len(all_cookies)} SharePoint page cookies from {self._current_host()}. Continuing…"
            )
            self.accept()
            return True

        if force_status:
            self._status.setText(
                f"{len(all_cookies)} cookies found, but no SharePoint cookies were scoped to the current page. "
                "Complete sign-in, wait for the SharePoint page to load, then click Check Session."
            )
        return False

    # ------------------------------------------------------------------ #
    #  Manual controls                                                     #
    # ------------------------------------------------------------------ #

    def _manual_check(self) -> None:
        """Explicit button press — mirrors CaptureSessionButton_OnClick."""
        _append_debug_log(f"manual_check_clicked | url={self._browser.url().toString()} | cookie_count={len(self._cookies)}")
        self._cookie_store.loadAllCookies()
        # Also re-read JS cookies from current page
        self._browser.page().runJavaScript("document.cookie", self._on_js_cookies)
        QTimer.singleShot(1200, lambda: self._try_capture(force_status=True))

    def _proceed_anyway(self) -> None:
        self._captured_cookies = list(self._cookies.values())
        _append_debug_log(
            f"proceed_anyway_clicked | url={self._browser.url().toString()} | cookie_count={len(self._captured_cookies)}"
        )
        if self._captured_cookies:
            self._status.setText(f"Proceeding with {len(self._captured_cookies)} cookies.")
            self.accept()
        else:
            self._status.setText("No cookies yet — please sign in first.")

    def captured_cookies(self) -> list[CookieSnapshot]:
        return list(self._captured_cookies)
