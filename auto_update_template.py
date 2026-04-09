from __future__ import annotations

import json
import os
import re
import shutil
import ssl
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class AutoUpdateConfig:
    # GitHub repo slug, e.g. "owner/repo"
    repo: str
    # Current app version string, e.g. "v1.2.3"
    current_version: str
    # App name used in download file naming
    app_name: str
    # Preferred asset stem. If None, uses executable stem.
    prefer_asset_stem: str | None = None
    # Storage directory for update payloads.
    data_dir: Path = Path.cwd() / "updates"
    # GitHub API user-agent
    user_agent: str = "github-auto-updater-template"
    # User preferences file. If None, uses data_dir/update_prefs.json.
    prefs_file: Path | None = None


@dataclass(frozen=True)
class UpdatePrefs:
    # User can turn update check on/off.
    check_on_startup: bool = True
    # Optional: user skips a specific version once.
    skipped_version: str = ""


@dataclass(frozen=True)
class UpdateInfo:
    has_update: bool
    latest_version: str
    download_url: str
    asset_name: str
    expected_size: int | None
    release_url: str


class GithubAutoUpdater:
    _SOURCE_ZIP_MARKERS = ("source code", "source_code", "原始碼", "源代码")

    def __init__(self, config: AutoUpdateConfig) -> None:
        self.cfg = config

    @staticmethod
    def _ssl_context() -> ssl.SSLContext | None:
        """Use certifi CA bundle when available to avoid Windows cert chain issues."""
        try:
            import certifi  # type: ignore

            return ssl.create_default_context(cafile=certifi.where())
        except Exception:
            return None

    def _urlopen(self, req: urllib.request.Request, timeout: int):
        ctx = self._ssl_context()
        if ctx is None:
            return urllib.request.urlopen(req, timeout=timeout)
        return urllib.request.urlopen(req, timeout=timeout, context=ctx)

    def _prefs_path(self) -> Path:
        if self.cfg.prefs_file is not None:
            return Path(self.cfg.prefs_file)
        return self.cfg.data_dir / "update_prefs.json"

    def load_prefs(self) -> UpdatePrefs:
        p = self._prefs_path()
        try:
            raw: dict[str, Any] = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            return UpdatePrefs()
        return UpdatePrefs(
            check_on_startup=bool(raw.get("check_on_startup", True)),
            skipped_version=str(raw.get("skipped_version") or "").strip(),
        )

    def save_prefs(self, prefs: UpdatePrefs) -> None:
        p = self._prefs_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(p.suffix + ".tmp")
        payload = {
            "check_on_startup": bool(prefs.check_on_startup),
            "skipped_version": str(prefs.skipped_version or "").strip(),
        }
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(p)

    def set_check_on_startup(self, enabled: bool) -> UpdatePrefs:
        prefs = self.load_prefs()
        updated = UpdatePrefs(check_on_startup=bool(enabled), skipped_version=prefs.skipped_version)
        self.save_prefs(updated)
        return updated

    def should_check_now(self, *, force: bool = False) -> bool:
        if force:
            return True
        return self.load_prefs().check_on_startup

    def mark_skip_version(self, version: str) -> UpdatePrefs:
        prefs = self.load_prefs()
        updated = UpdatePrefs(check_on_startup=prefs.check_on_startup, skipped_version=version.strip())
        self.save_prefs(updated)
        return updated

    def clear_skipped_version(self) -> UpdatePrefs:
        prefs = self.load_prefs()
        updated = UpdatePrefs(check_on_startup=prefs.check_on_startup, skipped_version="")
        self.save_prefs(updated)
        return updated

    @staticmethod
    def _version_key(v: str) -> tuple[int, ...]:
        nums = [int(x) for x in re.findall(r"\d+", (v or "").lower().lstrip("v"))]
        return tuple(nums) if nums else (0,)

    @staticmethod
    def _asset_size(a: dict) -> int | None:
        s = a.get("size")
        if isinstance(s, int):
            return s
        if isinstance(s, float):
            return int(s)
        return None

    @staticmethod
    def _is_source_zip(name_lower: str) -> bool:
        return any(m in name_lower for m in GithubAutoUpdater._SOURCE_ZIP_MARKERS)

    @staticmethod
    def _is_noise_zip(name_lower: str) -> bool:
        # You can expand this list for your own workflow.
        return "release_bundle" in name_lower or name_lower == "default.zip"

    def _pick_asset(
        self, data: dict, page_url: str, prefer_exe_name: str | None
    ) -> tuple[str, str, int | None]:
        assets = data.get("assets")
        rows: list[tuple[str, str, int | None]] = []
        if isinstance(assets, list):
            for a in assets:
                if not isinstance(a, dict):
                    continue
                u = str(a.get("browser_download_url") or "").strip()
                if not u:
                    continue
                name = str(a.get("name") or "").strip()
                if not name:
                    continue
                rows.append((u, name, self._asset_size(a)))

        prefer = (prefer_exe_name or "").lower().strip()
        prefer_stem = prefer[:-4] if prefer.endswith(".exe") else prefer
        prefer_zip = f"{prefer_stem}.zip" if prefer_stem else ""

        zip_rows = []
        for u, n, sz in rows:
            nl = n.lower()
            if nl.endswith(".zip") and not self._is_source_zip(nl) and not self._is_noise_zip(nl):
                zip_rows.append((u, n, sz))
        if prefer_zip:
            for u, n, sz in zip_rows:
                if n.lower() == prefer_zip:
                    return u, n, sz
        if zip_rows:
            return zip_rows[0]

        exe_rows = [(u, n, sz) for u, n, sz in rows if n.lower().endswith(".exe")]
        if prefer:
            for u, n, sz in exe_rows:
                if n.lower() == prefer:
                    return u, n, sz
        if exe_rows:
            return exe_rows[0]

        if rows:
            return rows[0]
        z = str(data.get("zipball_url") or "").strip()
        if z:
            return z, "source.zip", None
        return page_url, "release_page", None

    def check_latest(self) -> UpdateInfo:
        url = f"https://api.github.com/repos/{self.cfg.repo}/releases/latest"
        req = urllib.request.Request(
            url,
            headers={"Accept": "application/vnd.github+json", "User-Agent": self.cfg.user_agent},
        )
        with self._urlopen(req, timeout=15) as r:
            raw = r.read().decode("utf-8", errors="replace")
            data = json.loads(raw)

        latest = str(data.get("tag_name") or "").strip()
        page_url = str(data.get("html_url") or f"https://github.com/{self.cfg.repo}/releases").strip()

        prefer_stem = (self.cfg.prefer_asset_stem or "").strip()
        if not prefer_stem and getattr(sys, "frozen", False):
            prefer_stem = Path(sys.executable).stem
        prefer_exe_name = f"{prefer_stem}.exe" if prefer_stem else None

        dl, asset_name, sz = self._pick_asset(data, page_url, prefer_exe_name)
        has_update = self._version_key(latest) > self._version_key(self.cfg.current_version)
        return UpdateInfo(
            has_update=has_update,
            latest_version=latest,
            download_url=dl,
            asset_name=asset_name,
            expected_size=sz,
            release_url=page_url,
        )

    def check_latest_if_enabled(self, *, force: bool = False) -> UpdateInfo | None:
        """
        Return None when user disabled update checks.
        Use force=True for manual "Check updates now" action.
        """
        if not self.should_check_now(force=force):
            return None
        return self.check_latest()

    def should_offer_update(self, info: UpdateInfo) -> bool:
        """
        Whether UI should show update dialog:
        - has update
        - not skipped by user
        """
        if not info.has_update:
            return False
        prefs = self.load_prefs()
        if prefs.skipped_version and prefs.skipped_version == info.latest_version:
            return False
        return True

    def download(self, info: UpdateInfo) -> Path:
        self.cfg.data_dir.mkdir(parents=True, exist_ok=True)
        safe_ver = re.sub(r"[^0-9A-Za-z._-]+", "_", info.latest_version or "latest")
        ext = Path(urllib.parse.urlparse(info.download_url).path).suffix.lower()
        if ext not in {".zip", ".exe"}:
            ext = ".zip"
        out = self.cfg.data_dir / f"{self.cfg.app_name}_{safe_ver}{ext}"

        req = urllib.request.Request(info.download_url, headers={"User-Agent": self.cfg.user_agent})
        with self._urlopen(req, timeout=180) as r:
            with open(out, "wb") as f:
                shutil.copyfileobj(r, f)

        if not out.is_file():
            raise RuntimeError("download failed: file not found")
        got = out.stat().st_size
        if info.expected_size is not None and got != info.expected_size:
            out.unlink(missing_ok=True)
            raise RuntimeError(f"download size mismatch: got={got}, expected={info.expected_size}")

        if ext == ".zip":
            with zipfile.ZipFile(out, "r") as zf:
                infos = zf.infolist()
                if not infos:
                    out.unlink(missing_ok=True)
                    raise RuntimeError("invalid zip: empty archive")
                has_exe = any((not i.is_dir()) and i.filename.lower().endswith(".exe") for i in infos)
                if not has_exe:
                    out.unlink(missing_ok=True)
                    raise RuntimeError("invalid zip: no exe in archive")
        if ext == ".exe":
            with open(out, "rb") as f:
                if f.read(2) != b"MZ":
                    out.unlink(missing_ok=True)
                    raise RuntimeError("invalid exe header")
        return out

    def apply_update_without_restart(self, downloaded: Path) -> bool:
        """
        Replace current executable and exit app.
        - Old exe is renamed to *.bak
        - No auto restart (more stable on some machines)
        """
        if not getattr(sys, "frozen", False):
            raise RuntimeError("apply_update_without_restart only works in frozen exe mode")

        target_exe = Path(sys.executable).resolve()
        pid = os.getpid()

        def q(s: str) -> str:
            return s.replace("'", "''")

        ps = (
            f"$pidToWait={pid};"
            f"$target='{q(str(target_exe))}';"
            f"$new='{q(str(downloaded))}';"
            "$newExt=[IO.Path]::GetExtension($new).ToLowerInvariant();"
            "$tmp=Join-Path $env:TEMP ('autoupd_' + [guid]::NewGuid().ToString('N'));"
            "$staged=$target + '.new';"
            "$backup=$target + '.bak';"
            "for($i=0;$i -lt 90;$i++){"
            "  if(-not (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue)){break};"
            "  Start-Sleep -Seconds 1"
            "};"
            "try{"
            "  if($newExt -eq '.zip'){"
            "    Expand-Archive -LiteralPath $new -DestinationPath $tmp -Force;"
            "    $tn=[IO.Path]::GetFileName($target);"
            "    $exe=Get-ChildItem -LiteralPath $tmp -Recurse -Filter $tn -ErrorAction SilentlyContinue | Select-Object -First 1;"
            "    if(-not $exe){ $exe=Get-ChildItem -LiteralPath $tmp -Recurse -Filter '*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 };"
            "    if(-not $exe){throw 'zip_no_exe'};"
            "    Copy-Item -LiteralPath $exe.FullName -Destination $staged -Force"
            "  } else {"
            "    Copy-Item -LiteralPath $new -Destination $staged -Force"
            "  };"
            "  if(Test-Path -LiteralPath $backup){Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue};"
            "  Move-Item -LiteralPath $target -Destination $backup -Force;"
            "  Move-Item -LiteralPath $staged -Destination $target -Force;"
            "} catch {"
            "  if(Test-Path -LiteralPath $backup){"
            "    if(Test-Path -LiteralPath $target){Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue};"
            "    Move-Item -LiteralPath $backup -Destination $target -Force -ErrorAction SilentlyContinue"
            "  }"
            "} finally {"
            "  if(Test-Path -LiteralPath $tmp){Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue};"
            "  if(Test-Path -LiteralPath $new){Remove-Item -LiteralPath $new -Force -ErrorAction SilentlyContinue};"
            "  if(Test-Path -LiteralPath $staged){Remove-Item -LiteralPath $staged -Force -ErrorAction SilentlyContinue}"
            "}"
        )
        try:
            subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                close_fds=True,
            )
            return True
        except Exception:
            return False


def quick_start_example() -> None:
    """
    Example (replace fields):
    1) Fill repo/current_version/app_name.
    2) Call check_latest() -> download() -> apply_update_without_restart().
    """
    updater = GithubAutoUpdater(
        AutoUpdateConfig(
            repo="owner/repo",
            current_version="v1.0.0",
            app_name="my_app",
        )
    )
    info = updater.check_latest()
    if not info.has_update:
        print("Already latest.")
        return
    file_path = updater.download(info)
    print(f"Downloaded: {file_path}")
    print("If you are in frozen mode, call apply_update_without_restart(file_path).")

