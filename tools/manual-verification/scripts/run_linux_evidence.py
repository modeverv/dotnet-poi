#!/usr/bin/env python3
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


ROOT = Path("/workspace")
EVIDENCE_ROOT = ROOT / "tools" / "manual-verification" / "evidence"
WORK_ROOT_NAME = "workfiles"
IMAGE_ROOT_NAME = "images"
SESSION_LOG_NAME = "session.log"
INDEX_NAME = "INDEX.md"


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def now_jst():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))


def read_project_version():
    csproj = ROOT / "src" / "DotnetPoi.Core" / "DotnetPoi.Core.csproj"
    try:
        root = ET.parse(csproj).getroot()
        for elem in root.iter():
            if elem.tag.endswith("VersionPrefix") and elem.text:
                return elem.text.strip()
    except Exception:
        pass
    return os.environ.get("DOTNETPOI_VERSION", "unknown")


def run(cmd, check=True):
    result = subprocess.run(cmd, text=True, capture_output=True)
    if check and result.returncode != 0:
        raise RuntimeError(
            f"command failed: {' '.join(cmd)}\nstdout={result.stdout}\nstderr={result.stderr}"
        )
    return result


def git_revision():
    env_rev = os.environ.get("DOTNETPOI_REVISION")
    if env_rev:
        return env_rev
    result = run(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"], check=False)
    return result.stdout.strip() if result.returncode == 0 and result.stdout.strip() else "nogit"


def libreoffice_version():
    result = run(["libreoffice", "--version"])
    return result.stdout.strip()


def log(path, message):
    stamp = now_jst().strftime("%Y-%m-%d %H:%M:%S %z")
    line = f"[{stamp}] {message}"
    print(line, flush=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


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
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return desktop
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(f"could not connect to LibreOffice on port {port}: {last_error}")


def start_office(evidence_dir):
    profile = Path("/tmp/dotnet-poi-phase11-evidence-lo-profile")
    if profile.exists():
        shutil.rmtree(profile)
    profile_uri = uno.systemPathToFileUrl(str(profile))
    port = 23411
    cmd = [
        "libreoffice",
        f"-env:UserInstallation={profile_uri}",
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        "--nodefault",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    proc = subprocess.Popen(
        cmd,
        stdout=(evidence_dir / "libreoffice.stdout.log").open("w", encoding="utf-8"),
        stderr=(evidence_dir / "libreoffice.stderr.log").open("w", encoding="utf-8"),
        text=True,
    )
    return proc, connect_to_office(port)


def stop_office(proc, desktop):
    try:
        desktop.terminate()
    except Exception:
        pass
    try:
        proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait(timeout=10)


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
            "source": generated("manual-simple.xlsx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "xlsm",
            "source": generated("manual-simple.xlsm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted xlsx",
            "source": generated("manual-encrypted.xlsx"),
            "password": "f",
            "encrypted": True,
        },
        {
            "kind": "docx",
            "source": generated("manual-simple.docx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "docm",
            "source": generated("manual-simple.docm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted docx",
            "source": generated("manual-encrypted.docx"),
            "password": "f",
            "encrypted": True,
        },
        {
            "kind": "pptx",
            "source": generated("manual-simple.pptx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "pptm",
            "source": generated("manual-simple.pptm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted pptx",
            "source": generated("manual-encrypted.pptx"),
            "password": "f",
            "encrypted": True,
        },
    ]


def filter_for_png(kind):
    if "doc" in kind:
        return "writer_png_Export"
    if "ppt" in kind:
        return "impress_png_Export"
    return "calc_png_Export"


def safe_name(text):
    allowed = []
    for ch in text.lower().replace(" ", "-"):
        allowed.append(ch if ch.isalnum() or ch in "._-" else "-")
    return "".join(allowed).strip("-")


def load_document(desktop, path, password):
    args = [prop("Hidden", True), prop("ReadOnly", True), prop("MacroExecutionMode", 0)]
    if password:
        args.append(prop("Password", password))
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, tuple(args))


def export_png(doc, out_path, kind, password):
    args = [
        prop("FilterName", filter_for_png(kind)),
        prop("Overwrite", True),
    ]
    if password:
        args.append(prop("Password", password))
    doc.storeToURL(uno.systemPathToFileUrl(str(out_path)), tuple(args))


def close_document(doc):
    try:
        doc.close(True)
    except Exception:
        doc.dispose()


def verify_case(desktop, case, ordinal, evidence_dir, work_dir, image_dir, session_log):
    source_rel = case["source"]
    result = {
        "kind": case["kind"],
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
    doc = None
    reopened = None
    try:
        log(session_log, f"case start: {case['kind']} {source_rel}")
        doc = load_document(desktop, target, case["password"])
        if doc is None:
            raise RuntimeError("LibreOffice returned no document")
        result["open"] = "PASS"
        image_name = f"{ordinal:02d}-{safe_name(case['kind'])}-{safe_name(source.stem)}.png"
        image_path = image_dir / image_name
        export_png(doc, image_path, case["kind"], case["password"])
        result["png"] = f"{IMAGE_ROOT_NAME}/{image_name}"
        close_document(doc)
        doc = None

        reopened = load_document(desktop, target, case["password"])
        if reopened is None:
            raise RuntimeError("LibreOffice returned no document on reopen")
        result["reopen"] = "PASS"
        result["status"] = "PASS"
    except Exception as exc:
        result["status"] = "FAIL"
        result["notes"].append(f"{exc.__class__.__name__}: {exc}")
        log(session_log, f"case failed: {case['kind']} {source_rel}: {exc}")
    finally:
        if doc is not None:
            close_document(doc)
        if reopened is not None:
            close_document(reopened)
    log(session_log, f"case result: {result}")
    return result


def table_rows(results):
    rows = []
    for r in results:
        image = f'<img src="{r["png"]}" width="320" alt="{r["kind"]} evidence">' if r["png"] else ""
        notes = "<br/>".join(r["notes"])
        rows.append(
            f"| {r['kind']} | {r['source']} | {r['encrypted']} | {r['open']} | "
            f"{r['reopen']} | {r['status']} | {image} | {notes} |"
        )
    return "\n".join(rows)


def write_index(evidence_dir, start, end, version, revision, lo_version, results):
    passing = sum(1 for r in results if r["status"] == "PASS")
    missing = sum(1 for r in results if r["status"] == "MISSING")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    overall = "PASS_WITH_MISSING_FIXTURES" if failed == 0 and missing else ("PASS" if failed == 0 else "FAIL")
    content = f"""# DotnetPOI {evidence_dir.name} Linux LibreOffice Evidence

- Project version: `{version}`
- Git revision: `{revision}`
- Captured: `{start.strftime('%Y-%m-%d %H:%M:%S %z')}` - `{end.strftime('%Y-%m-%d %H:%M:%S %z')}`
- Environment: Docker service `libreoffice`, container `dotnet-poi-phase11-libreoffice`
- LibreOffice: `{lo_version}`
- Source root: `tools/manual-verification/generated-documents`
- Overall: `{overall}`
- Result counts: `{passing}` pass, `{missing}` missing fixture, `{failed}` fail

This evidence pass opens each available file through LibreOffice UNO, rejects failures/exceptions as a failed case, reopens the work copy, exports a PNG preview, and writes this index for GitHub review.

## Matrix

| kind | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---:|---:|---:|---:|---|---|
{table_rows(results)}

## Notes

- `MISSING` means generated manual documents are not present; run `dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj`.
- The original files are not modified; work copies are written under `{WORK_ROOT_NAME}/`.
- PNG previews are exported by LibreOffice itself, not by the browser screenshot path.
- Password for generated encrypted files: `f`.
"""
    (evidence_dir / INDEX_NAME).write_text(content, encoding="utf-8")
    return overall


def main():
    version = os.environ.get("DOTNETPOI_VERSION") or read_project_version()
    revision = git_revision()
    evidence_id = os.environ.get("DOTNETPOI_EVIDENCE_ID") or f"v{version}-{revision}"
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
    lo_version = libreoffice_version()
    log(session_log, f"version={version} revision={revision} libreoffice={lo_version}")

    proc, desktop = start_office(evidence_dir)
    results = []
    try:
        for ordinal, case in enumerate(case_matrix(), start=1):
            results.append(verify_case(desktop, case, ordinal, evidence_dir, work_dir, image_dir, session_log))
    finally:
        stop_office(proc, desktop)

    end = now_jst()
    overall = write_index(evidence_dir, start, end, version, revision, lo_version, results)
    log(session_log, f"index written: {evidence_dir / INDEX_NAME}")
    log(session_log, f"overall={overall}")
    return 1 if any(r["status"] == "FAIL" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
