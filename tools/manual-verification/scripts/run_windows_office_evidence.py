#!/usr/bin/env python3
"""
Phase 11 — Windows Office manual verification evidence runner.

Opens each fixture in the corresponding Microsoft Office COM application,
captures a screenshot, closes, reopens, then writes an INDEX.md evidence file.

Requirements (Windows only):
    pip install pywin32 Pillow

Run via run-windows-office-evidence.bat, or directly:
    python run_windows_office_evidence.py

Environment variables (optional — auto-detected if not set):
    DOTNETPOI_VERSION     project version (read from DotnetPoi.Core.csproj)
    DOTNETPOI_REVISION    git short SHA
    DOTNETPOI_EVIDENCE_ID output directory name under tools/manual-verification/evidence/
"""
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
EVIDENCE_ROOT = ROOT / "tools" / "manual-verification" / "evidence"
WORK_ROOT_NAME = "workfiles"
IMAGE_ROOT_NAME = "images"
SESSION_LOG_NAME = "session.log"
INDEX_NAME = "INDEX.md"

# msoAutomationSecurityLow — suppresses macro / security prompts when opening via COM.
# This is intentional for compatibility verification: the goal is to confirm that files
# open without repair dialogs, not to test macro execution.
MSO_AUTOMATION_SECURITY_LOW = 1


def now_jst():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))


def read_project_version():
    csproj = ROOT / "src" / "DotnetPoi.Common" / "DotnetPoi.Common.csproj"
    try:
        root = ET.parse(csproj).getroot()
        for elem in root.iter():
            if elem.tag.endswith("VersionPrefix") and elem.text:
                return elem.text.strip()
    except Exception:
        pass
    return os.environ.get("DOTNETPOI_VERSION", "unknown")


def git_revision():
    env_rev = os.environ.get("DOTNETPOI_REVISION")
    if env_rev:
        return env_rev
    result = subprocess.run(
        ["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"],
        text=True, capture_output=True,
    )
    return result.stdout.strip() if result.returncode == 0 and result.stdout.strip() else "nogit"


def log(path, message):
    stamp = now_jst().strftime("%Y-%m-%d %H:%M:%S %z")
    line = f"[{stamp}] {message}"
    print(line, flush=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def existing(*candidates):
    for candidate in candidates:
        path = ROOT / candidate
        if path.exists():
            return candidate
    return None


def generated(name):
    return existing(f"tools/manual-verification/generated-documents/{name}")


def case_matrix():
    return [
        {
            "kind": "xlsx",
            "app": "excel",
            "source": generated("manual-simple.xlsx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "xlsm",
            "app": "excel",
            "source": generated("manual-simple.xlsm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted xlsx",
            "app": "excel",
            "source": generated("manual-encrypted.xlsx"),
            "password": "f",
            "encrypted": True,
        },
        {
            "kind": "docx",
            "app": "word",
            "source": generated("manual-simple.docx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "docm",
            "app": "word",
            "source": generated("manual-simple.docm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted docx",
            "app": "word",
            "source": generated("manual-encrypted.docx"),
            "password": "f",
            "encrypted": True,
        },
#       {
#            "kind": "pptx",
#            "app": "powerpoint",
#            "source": generated("manual-simple.pptx"),
#            "password": None,
#            "encrypted": False,
#        },
#        {
#            "kind": "pptm",
#            "app": "powerpoint",
#            "source": generated("manual-simple.pptm"),
#            "password": None,
#            "encrypted": False,
#        },
#         {
#              "kind": "encrypted pptx",
#              "app": "powerpoint",
#              "source": generated("manual-encrypted.pptx"),
#              "password": "f",
#              "encrypted": True,
#        },
        {
             "kind": "xls",
             "app": "excel",
             "source": generated("manual-simple.xls"),
             "password": None,
             "encrypted": False,
        },
        {
            "kind": "doc",
            "app": "word",
            "source": generated("manual-simple.doc"),
            "password": None,
            "encrypted": False,
        },
#        {
#           "kind": "ppt",
#           "app": "powerpoint",
#           "source": generated("manual-simple.ppt"),
#           "password": None,
#           "encrypted": False,
#       },
    ]


def safe_name(text):

    allowed = []
    for ch in text.lower().replace(" ", "-"):
        allowed.append(ch if ch.isalnum() or ch in "._-" else "-")
    return "".join(allowed).strip("-")


# ─── prerequisites ─────────────────────────────────────────────────────────────

def check_prerequisites():
    errors = []
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        errors.append("pywin32 is not installed — run: pip install pywin32")
    try:
        from PIL import ImageGrab  # noqa: F401
    except ImportError:
        errors.append("Pillow is not installed — run: pip install Pillow")
    if errors:
        for e in errors:
            print(f"ERROR: {e}", file=sys.stderr)
        raise SystemExit(3)


# ─── Office app helpers ────────────────────────────────────────────────────────

def _try_bring_to_front(hwnd_or_none):
    try:
        import win32gui
        import win32con
        if hwnd_or_none:
            win32gui.ShowWindow(hwnd_or_none, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd_or_none)
    except Exception:
        pass


def _graceful_quit(prog_id):
    """Quit a running Office app via COM without forcefully killing it."""
    try:
        import win32com.client
        app = win32com.client.GetActiveObject(prog_id)
        try:
            app.Quit()
        except Exception:
            pass
        del app
    except Exception:
        pass


def _force_kill(exe_name):
    subprocess.run(
        ["taskkill", "/F", "/IM", exe_name],
        capture_output=True, text=True,
    )


def close_all_office():
    """Gracefully quit all three Office apps, then force-kill any stragglers."""
    for prog_id in ("Excel.Application", "Word.Application", "PowerPoint.Application"):
        _graceful_quit(prog_id)
    time.sleep(2)
    # Force-kill anything that didn't respond to Quit()
    for exe in ("EXCEL.EXE", "WINWORD.EXE", "POWERPNT.EXE"):
        _force_kill(exe)
    time.sleep(1)


def app_versions():
    """Return {excel, word, powerpoint} version strings via COM DispatchEx."""
    import win32com.client
    versions = {}
    for key, prog_id in [
        ("excel", "Excel.Application"),
        ("word", "Word.Application"),
        ("powerpoint", "PowerPoint.Application"),
    ]:
        try:
            app = win32com.client.DispatchEx(prog_id)
            versions[key] = getattr(app, "Version", "unknown")
            try:
                app.Quit()
            except Exception:
                pass
            del app
            time.sleep(1)
        except Exception as exc:
            versions[key] = f"error: {exc}"
    return versions


# ─── Open / close per app ──────────────────────────────────────────────────────

def open_excel(path, password):
    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    if password:
        wb = excel.Workbooks.Open(str(path), 0, False, 5, password)
    else:
        wb = excel.Workbooks.Open(str(path))
    try:
        _try_bring_to_front(excel.Hwnd)
    except Exception:
        pass
    return excel, wb


def close_excel(excel, wb):
    try:
        if wb is not None:
            wb.Close(False)
    except Exception:
        pass
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass


def open_word(path, password):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True
    try:
        word.DisplayAlerts = False
    except Exception:
        pass
    word.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    if password:
        doc = word.Documents.Open(
            FileName=str(path),
            PasswordDocument=password,
            AddToRecentFiles=False,
        )
    else:
        doc = word.Documents.Open(FileName=str(path), AddToRecentFiles=False)
    try:
        word.Activate()
    except Exception:
        pass
    return word, doc


def close_word(word, doc):
    try:
        if doc is not None:
            doc.Close(0)  # wdDoNotSaveChanges = 0
    except Exception:
        pass
    try:
        if word is not None:
            word.Quit()
    except Exception:
        pass


def open_powerpoint(path, password):
    import win32com.client
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        ppt.Visible = True
    except Exception:
        pass
    ppt.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    if password:
        pres = ppt.Presentations.Open(
            FileName=str(path),
            Password=password,
            ReadOnly=False,
            Untitled=False,
        )
    else:
        pres = ppt.Presentations.Open(
            FileName=str(path),
            ReadOnly=False,
            Untitled=False,
        )
    try:
        ppt.Activate()
    except Exception:
        pass
    return ppt, pres


def close_powerpoint(ppt, pres):
    try:
        if pres is not None:
            pres.Close()
    except Exception:
        pass
    try:
        if ppt is not None:
            ppt.Quit()
    except Exception:
        pass


def open_document(case, path):
    app_key = case["app"]
    password = case["password"] or None
    if app_key == "excel":
        return open_excel(path, password)
    elif app_key == "word":
        return open_word(path, password)
    else:
        return open_powerpoint(path, password)


def close_document(case, app, doc):
    app_key = case["app"]
    if app_key == "excel":
        close_excel(app, doc)
    elif app_key == "word":
        close_word(app, doc)
    else:
        close_powerpoint(app, doc)


# ─── screenshot ───────────────────────────────────────────────────────────────

def capture_screenshot(image_path, delay=1.5):
    time.sleep(delay)
    from PIL import ImageGrab
    img = ImageGrab.grab()
    img.save(str(image_path), "PNG")


# ─── verify_case ──────────────────────────────────────────────────────────────

def verify_case(case, ordinal, work_dir, image_dir, session_log):
    source_rel = case["source"]
    result = {
        "kind": case["kind"],
        "app": case["app"],
        "source": source_rel or "",
        "encrypted": "yes" if case["encrypted"] else "no",
        "open": "MISSING",
        "reopen": "MISSING",
        "png": "",
        "status": "MISSING",
        "notes": [],
    }
    if not source_rel:
        result["notes"].append("No fixture is currently available for this category.")
        return result

    source = ROOT / source_rel
    target = work_dir / source.name
    shutil.copy2(source, target)
    image_name = f"{ordinal:02d}-{safe_name(case['kind'])}-{safe_name(source.stem)}.png"
    image_path = image_dir / image_name

    try:
        log(session_log, f"case start: {case['kind']} {source_rel}")
        close_all_office()

        # open pass
        app_obj, doc_obj = None, None
        try:
            app_obj, doc_obj = open_document(case, target)
            result["open"] = "PASS"
            capture_screenshot(image_path)
            result["png"] = f"{IMAGE_ROOT_NAME}/{image_name}"
        finally:
            close_document(case, app_obj, doc_obj)

        time.sleep(2)
        close_all_office()

        # reopen pass
        app_obj, doc_obj = None, None
        try:
            app_obj, doc_obj = open_document(case, target)
            result["reopen"] = "PASS"
        finally:
            close_document(case, app_obj, doc_obj)

        result["status"] = "PASS"
    except Exception as exc:
        result["status"] = "FAIL"
        result["notes"].append(f"{exc.__class__.__name__}: {exc}")
        log(session_log, f"case failed: {case['kind']} {source_rel}: {exc}")
    finally:
        close_all_office()

    log(session_log, f"case result: {result}")
    return result


# ─── report ───────────────────────────────────────────────────────────────────

def table_rows(results):
    rows = []
    for r in results:
        image = (
            f'<img src="{r["png"]}" width="320" alt="{r["kind"]} Windows Office evidence">'
            if r["png"] else ""
        )
        notes = "<br/>".join(r["notes"])
        rows.append(
            f"| {r['kind']} | {r['app']} | {r['source']} | {r['encrypted']} | "
            f"{r['open']} | {r['reopen']} | {r['status']} | {image} | {notes} |"
        )
    return "\n".join(rows)


def write_index(evidence_dir, start, end, version, revision, versions, results):
    passing = sum(1 for r in results if r["status"] == "PASS")
    missing = sum(1 for r in results if r["status"] == "MISSING")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    overall = "FAIL" if failed else ("PASS_WITH_MISSING_FIXTURES" if missing else "PASS")
    content = f"""# DotnetPOI {evidence_dir.name} Windows Office Evidence

- Project version: `{version}`
- Git revision: `{revision}`
- Captured: `{start.strftime('%Y-%m-%d %H:%M:%S %z')}` - `{end.strftime('%Y-%m-%d %H:%M:%S %z')}`
- Environment: Windows Microsoft Office (COM automation)
- Excel: `{versions.get('excel', 'unknown')}`
- Word: `{versions.get('word', 'unknown')}`
- PowerPoint: `{versions.get('powerpoint', 'unknown')}`
- Source root: `tools/manual-verification/generated-documents`
- Overall: `{overall}`
- Result counts: `{passing}` pass, `{missing}` missing fixture, `{failed}` fail

Macro security: `AutomationSecurity = msoAutomationSecurityLow (1)` is applied on each COM
application before opening files. Macro prompts are suppressed by design for this compatibility
verification; the goal is to confirm files open without repair dialogs, not to test macro execution.

## Matrix

| kind | app | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---|---:|---:|---:|---:|---|---|
{table_rows(results)}

## Notes

- `MISSING` means generated manual documents are not present; run `dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj`.
- Original files are not modified; work copies are written under `{WORK_ROOT_NAME}/`.
- Password for generated encrypted files: `f`.
- Screenshots captured with `PIL.ImageGrab.grab()` (full-screen). Requires `pip install Pillow`.
- COM automation requires Microsoft Office to be installed and licensed on the host machine.
- `DispatchEx` is used per case to ensure a fresh process rather than attaching to an existing instance.
"""
    (evidence_dir / INDEX_NAME).write_text(content, encoding="utf-8")
    return overall


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    if sys.platform != "win32":
        print("ERROR: This script must be run on Windows.", file=sys.stderr)
        return 3

    check_prerequisites()

    version = os.environ.get("DOTNETPOI_VERSION") or read_project_version()
    revision = os.environ.get("DOTNETPOI_REVISION") or git_revision()
    evidence_id = os.environ.get("DOTNETPOI_EVIDENCE_ID") or f"v{version}-{revision}-windows"
    evidence_dir = EVIDENCE_ROOT / evidence_id
    work_dir = evidence_dir / WORK_ROOT_NAME
    image_dir = evidence_dir / IMAGE_ROOT_NAME
    if evidence_dir.exists():
        shutil.rmtree(evidence_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    image_dir.mkdir(parents=True, exist_ok=True)
    session_log = evidence_dir / SESSION_LOG_NAME
    session_log.touch()

    start = now_jst()
    log(session_log, "collecting Office versions...")
    versions = app_versions()
    log(session_log, f"version={version} revision={revision} office={versions}")

    results = []
    try:
        for ordinal, case in enumerate(case_matrix(), start=1):
            results.append(verify_case(case, ordinal, work_dir, image_dir, session_log))
    finally:
        close_all_office()

    end = now_jst()
    overall = write_index(evidence_dir, start, end, version, revision, versions, results)
    log(session_log, f"index written: {evidence_dir / INDEX_NAME}")
    log(session_log, f"overall={overall}")
    return 1 if any(r["status"] == "FAIL" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
