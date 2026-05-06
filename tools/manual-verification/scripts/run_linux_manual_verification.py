#!/usr/bin/env python3
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


ROOT = Path("/workspace")
SOURCE_DIR = ROOT / "tools" / "manual-verification" / "generated-documents"
EVIDENCE = ROOT / "tools" / "manual-verification" / "evidence" / "linux"
WORKFILES = EVIDENCE / "workfiles"
SESSION_LOG = EVIDENCE / "session.log"
SUMMARY = EVIDENCE / "summary.md"

CASE_FILES = [
    ("manual-simple.xlsx", None),
    ("manual-simple.xlsm", None),
    ("manual-encrypted.xlsx", "f"),
    ("manual-simple.docx", None),
    ("manual-simple.docm", None),
    ("manual-encrypted.docx", "f"),
    ("manual-simple.pptx", None),
    ("manual-simple.pptm", None),
    ("manual-encrypted.pptx", "f"),
]


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def now_jst():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))


def log(message):
    stamp = now_jst().strftime("%Y-%m-%d %H:%M:%S %z")
    line = f"[{stamp}] {message}"
    print(line, flush=True)
    with SESSION_LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(cmd, check=True):
    log("$ " + " ".join(cmd))
    result = subprocess.run(cmd, text=True, capture_output=True)
    if result.stdout:
        for line in result.stdout.rstrip().splitlines():
            log("stdout: " + line)
    if result.stderr:
        for line in result.stderr.rstrip().splitlines():
            log("stderr: " + line)
    if check and result.returncode != 0:
        raise RuntimeError(f"command failed: {' '.join(cmd)} rc={result.returncode}")
    return result


def libreoffice_version():
    result = subprocess.run(["libreoffice", "--version"], text=True, capture_output=True, check=True)
    return result.stdout.strip()


def connect_to_office(port, timeout=45):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            ctx = resolver.resolve(url)
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return desktop
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(f"could not connect to LibreOffice on port {port}: {last_error}")


def start_office():
    profile = Path("/tmp/dotnet-poi-phase11-lo-profile")
    if profile.exists():
        shutil.rmtree(profile)
    profile_uri = uno.systemPathToFileUrl(str(profile))
    port = 23211
    cmd = [
        "libreoffice",
        f"-env:UserInstallation={profile_uri}",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        "--nodefault",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    log("$ " + " ".join(cmd))
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    desktop = connect_to_office(port)
    return proc, desktop


def stop_office(proc, desktop):
    try:
        desktop.terminate()
    except Exception as exc:
        log(f"LibreOffice terminate warning: {exc}")
    try:
        proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        log("LibreOffice did not terminate in time; killing process")
        proc.kill()
        proc.wait(timeout=10)
    stdout, stderr = proc.communicate()
    if stdout:
        for line in stdout.rstrip().splitlines():
            log("libreoffice stdout: " + line)
    if stderr:
        for line in stderr.rstrip().splitlines():
            log("libreoffice stderr: " + line)


def discover_cases():
    if not SOURCE_DIR.exists():
        raise RuntimeError(f"source directory does not exist: {SOURCE_DIR}")
    cases = []
    skipped = []
    for name, password in CASE_FILES:
        path = SOURCE_DIR / name
        if not path.exists():
            skipped.append((name, "generated manual document missing"))
            continue
        cases.append({"path": path, "password": password})
    return cases, skipped


def load_document(desktop, path, hidden, password):
    url = uno.systemPathToFileUrl(str(path))
    args = [
        prop("Hidden", hidden),
        prop("ReadOnly", False),
        prop("MacroExecutionMode", 0),
    ]
    if password:
        args.append(prop("Password", password))
    return desktop.loadComponentFromURL(url, "_blank", 0, tuple(args))


def close_document(doc):
    try:
        doc.close(True)
    except Exception:
        doc.dispose()


def wait_for_visible_window():
    # The linuxserver.io web desktop exposes the GUI through its own compositor.
    # UNO open/store/reopen is the compatibility signal here; visual inspection
    # is done separately through http://localhost:3110 when needed.
    time.sleep(1.5)


def verify_file(desktop, case):
    source = case["path"]
    password = case["password"]
    target = WORKFILES / source.name
    shutil.copy2(source, target)
    result = {
        "file": source.name,
        "kind": source.suffix.lower().lstrip("."),
        "open": "FAIL",
        "store": "FAIL",
        "reopen": "FAIL",
        "no_exception": "FAIL",
        "notes": [],
    }
    doc = None
    try:
        log(f"case start: {source.relative_to(ROOT)}")
        doc = load_document(desktop, target, hidden=False, password=password)
        if doc is None:
            raise RuntimeError("LibreOffice returned no document")
        result["open"] = "PASS"
        wait_for_visible_window()
        doc.store()
        result["store"] = "PASS"
        close_document(doc)
        doc = None

        doc = load_document(desktop, target, hidden=True, password=password)
        if doc is None:
            raise RuntimeError("LibreOffice returned no document on reopen")
        result["reopen"] = "PASS"
        result["no_exception"] = "PASS"
    except Exception as exc:
        result["notes"].append(f"{exc.__class__.__name__}: {exc}")
        log(f"case failed: {source.name}: {exc}")
    finally:
        if doc is not None:
            close_document(doc)
    log(f"case result: {result}")
    return result


def markdown_table(results):
    rows = [
        "| file | kind | open | store | reopen | no-exception | notes |",
        "|---|---|---:|---:|---:|---:|---|",
    ]
    for r in results:
        notes = "<br/>".join(r["notes"])
        rows.append(
            f"| {r['file']} | {r['kind']} | {r['open']} | {r['store']} | "
            f"{r['reopen']} | {r['no_exception']} | {notes} |"
        )
    return "\n".join(rows)


def write_summary(start, end, lo_version, results, skipped, overall):
    skipped_lines = "\n".join(f"- `{name}`: {reason}" for name, reason in skipped) or "- none"
    content = f"""# Linux LibreOffice Manual Verification Summary

- Started: {start.strftime('%Y-%m-%d %H:%M:%S %z')}
- Finished: {end.strftime('%Y-%m-%d %H:%M:%S %z')}
- Source files: `tools/manual-verification/generated-documents`
- Work files: `tools/manual-verification/evidence/linux/workfiles`
- LibreOffice: {lo_version}
- Container: `dotnet-poi-phase11-libreoffice`
- Web desktop: http://localhost:3110
- Overall: {overall}

## Results

{markdown_table(results)}

## Skipped

{skipped_lines}

## Notes

- This is an automated assist for Phase 11 manual verification: each work file is opened in LibreOffice, stored, closed, and reopened.
- The original files in `tools/manual-verification/generated-documents` are not modified.
- Encrypted generated files are opened with password `f`.
- Screenshot capture is not enabled in the current container image; visual inspection can be done through the web desktop.
"""
    SUMMARY.write_text(content, encoding="utf-8")


def main():
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    WORKFILES.mkdir(parents=True, exist_ok=True)
    SESSION_LOG.write_text("", encoding="utf-8")

    start = now_jst()
    lo_version = libreoffice_version()
    log(f"LibreOffice version: {lo_version}")
    log(f"source directory: {SOURCE_DIR}")
    log(f"evidence directory: {EVIDENCE}")

    cases, skipped = discover_cases()
    if not cases:
        raise RuntimeError(f"no supported Office files found under {SOURCE_DIR}")

    proc, desktop = start_office()
    results = []
    try:
        for source in cases:
            results.append(verify_file(desktop, source))
    finally:
        stop_office(proc, desktop)

    failures = [
        r["file"]
        for r in results
        if any(r[key] != "PASS" for key in ("open", "store", "reopen", "no_exception"))
    ]
    overall = "PASS" if not failures else "FAIL"
    end = now_jst()
    write_summary(start, end, lo_version, results, skipped, overall)
    log(f"summary written: {SUMMARY.relative_to(ROOT)}")
    log(f"overall result: {overall}")
    return 0 if overall == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
