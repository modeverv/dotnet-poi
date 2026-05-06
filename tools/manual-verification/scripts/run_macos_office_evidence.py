#!/usr/bin/env python3
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


def run(cmd, check=True, timeout=None):
    result = subprocess.run(cmd, text=True, capture_output=True, timeout=timeout)
    if check and result.returncode != 0:
        raise RuntimeError(
            f"command failed: {' '.join(cmd)}\nstdout={result.stdout}\nstderr={result.stderr}"
        )
    return result


def osascript(script, timeout=60):
    return run(["osascript", "-e", script], timeout=timeout)


def osascript_lines(lines, timeout=60):
    cmd = ["osascript"]
    for line in lines:
        cmd.extend(["-e", line])
    return run(cmd, timeout=timeout)


def git_revision():
    env_rev = os.environ.get("DOTNETPOI_REVISION")
    if env_rev:
        return env_rev
    result = run(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"], check=False)
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
        {
            "kind": "pptx",
            "app": "powerpoint",
            "source": generated("manual-simple.pptx"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "pptm",
            "app": "powerpoint",
            "source": generated("manual-simple.pptm"),
            "password": None,
            "encrypted": False,
        },
        {
            "kind": "encrypted pptx",
            "app": "powerpoint",
            "source": generated("manual-encrypted.pptx"),
            "password": "f",
            "encrypted": True,
        },
    ]


def safe_name(text):
    allowed = []
    for ch in text.lower().replace(" ", "-"):
        allowed.append(ch if ch.isalnum() or ch in "._-" else "-")
    return "".join(allowed).strip("-")


def preview_image(source_path, image_path):
    tmp_dir = image_path.parent / ".qlmanage"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True, exist_ok=True)
    result = run(
        ["qlmanage", "-t", "-s", "900", "-o", str(tmp_dir), str(source_path)],
        check=False,
        timeout=60,
    )
    produced = tmp_dir / f"{source_path.name}.png"
    if result.returncode != 0 or not produced.exists():
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return False, (result.stderr or result.stdout or "Quick Look did not produce a thumbnail").strip()
    shutil.copy2(produced, image_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return True, ""


def capture_screenshot(image_path):
    time.sleep(1.2)
    run(["/usr/sbin/screencapture", "-x", str(image_path)], timeout=20)


def check_screencapture_permission(evidence_dir):
    target = evidence_dir / "screencapture-preflight.png"
    if target.exists():
        target.unlink()
    result = run(["/usr/sbin/screencapture", "-x", str(target)], check=False, timeout=20)
    ok = result.returncode == 0 and target.exists()
    return ok, result


def run_applescript(script, timeout=60):
    return subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
        timeout=timeout,
    )


def is_macro_format(case):
    source = case.get("source") or ""
    return source.endswith((".xlsm", ".docm", ".pptm"))


def dismiss_macro_dialog(app_name, poll_attempts=6, poll_interval=1.5):
    """Click 'Enable Macros' if the macro security dialog appears after opening a macro-enabled file.

    Polls for the dialog button up to poll_attempts times with poll_interval seconds between each
    attempt. Safe to call when no dialog is present — all inner AppleScript blocks use try/end try.
    """
    button_list = '{"Enable Macros", "Enable Content", "マクロを有効にする", "コンテンツを有効にする"}'
    script = f"""\
tell application "{app_name}" to activate
delay 1
tell application "System Events"
    tell process "{app_name}"
        set dismissed to false
        repeat {poll_attempts} times
            try
                set allWindows to every window
                repeat with w in allWindows
                    try
                        set allButtons to every button of w
                        repeat with b in allButtons
                            try
                                if (name of b) is in {button_list} then
                                    click b
                                    set dismissed to true
                                end if
                            end try
                        end repeat
                    end try
                    if dismissed then exit repeat
                end repeat
            end try
            if dismissed then exit repeat
            delay {poll_interval}
        end repeat
    end tell
end tell
"""
    run_applescript(script, timeout=int(poll_attempts * poll_interval + 15))


def open_with_launch_services(app_name, path):
    run(["/usr/bin/open", "-n", "-a", app_name, path], timeout=30)
    time.sleep(4)


def type_password(app_name, password):
    process_name = app_name
    osascript_lines(
        [
            f'tell application "{app_name}" to activate',
            "delay 1",
            'tell application "System Events"',
            f'tell process "{process_name}" to set frontmost to true',
            "delay 1",
            f'keystroke "{password}"',
            "delay 0.2",
            'key code 36',
            'end tell',
        ],
        timeout=20,
    )
    time.sleep(6)


def app_versions():
    versions = {}
    for key, plist in [
        ("excel", "/Applications/Microsoft Excel.app/Contents/Info.plist"),
        ("word", "/Applications/Microsoft Word.app/Contents/Info.plist"),
        ("powerpoint", "/Applications/Microsoft PowerPoint.app/Contents/Info.plist"),
    ]:
        result = run(
            ["/usr/libexec/PlistBuddy", "-c", "Print :CFBundleShortVersionString", plist],
            check=False,
            timeout=10,
        )
        versions[key] = result.stdout.strip() if result.returncode == 0 else "unknown"
    return versions


def close_all():
    # Use separate best-effort calls. Office AppleScript dictionaries differ
    # enough that a single multi-app script can fail to parse before `try` runs.
    for app in ("Microsoft Excel", "Microsoft Word", "Microsoft PowerPoint"):
        subprocess.run(
            ["osascript", "-e", f'tell application "{app}" to quit saving no'],
            text=True,
            capture_output=True,
            timeout=20,
        )
    time.sleep(1)


def open_excel(path, password):
    if password:
        return []
    return []


def reopen_excel(path, password):
    return []


def open_word(path, password):
    # Password-protected Word fixtures are not yet available in this repo.
    return []


def reopen_word(path, password):
    return []


def open_powerpoint(path, password):
    # Password-protected PowerPoint fixtures are not yet available in this repo.
    return []


def reopen_powerpoint(path, password):
    return []


def open_lines(case, path):
    if case["app"] == "excel":
        return open_excel(path, case["password"])
    if case["app"] == "word":
        return open_word(path, case["password"])
    return open_powerpoint(path, case["password"])


def reopen_lines(case, path):
    if case["app"] == "excel":
        return reopen_excel(path, case["password"])
    if case["app"] == "word":
        return reopen_word(path, case["password"])
    return reopen_powerpoint(path, case["password"])


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
        app_name = {
            "excel": "Microsoft Excel",
            "word": "Microsoft Word",
            "powerpoint": "Microsoft PowerPoint",
        }[case["app"]]
        close_all()
        lines = open_lines(case, str(target))
        if lines:
            osascript_lines(lines, timeout=90)
        else:
            open_with_launch_services(app_name, str(target))
            if case["password"]:
                type_password(app_name, case["password"])
        if is_macro_format(case):
            dismiss_macro_dialog(app_name)
        result["open"] = "PASS"
        capture_screenshot(image_path)
        result["png"] = f"{IMAGE_ROOT_NAME}/{image_name}"
        close_all()
        lines = reopen_lines(case, str(target))
        if lines:
            osascript_lines(lines, timeout=90)
        else:
            open_with_launch_services(app_name, str(target))
            if case["password"]:
                type_password(app_name, case["password"])
        if is_macro_format(case):
            dismiss_macro_dialog(app_name)
        close_all()
        result["reopen"] = "PASS"
        result["status"] = "PASS"
    except Exception as exc:
        message = str(exc)
        if "キー操作の送信は許可されません" in message or "not allowed assistive access" in message:
            result["status"] = "PERMISSION_REQUIRED"
            result["notes"].append("Password dialog automation requires Accessibility/Input Monitoring permission. Run permission/bootstrap mode.")
        else:
            result["status"] = "FAIL"
        result["notes"].append(f"{exc.__class__.__name__}: {exc}")
        log(session_log, f"case failed: {case['kind']} {source_rel}: {exc}")
    finally:
        close_all()
    log(session_log, f"case result: {result}")
    return result


def table_rows(results):
    rows = []
    for r in results:
        image = f'<img src="{r["png"]}" width="320" alt="{r["kind"]} macOS Office evidence">' if r["png"] else ""
        notes = "<br/>".join(r["notes"])
        rows.append(
            f"| {r['kind']} | {r['app']} | {r['source']} | {r['encrypted']} | "
            f"{r['open']} | {r['reopen']} | {r['status']} | {image} | {notes} |"
        )
    return "\n".join(rows)


def write_index(evidence_dir, start, end, version, revision, versions, results):
    passing = sum(1 for r in results if r["status"] == "PASS")
    missing = sum(1 for r in results if r["status"] == "MISSING")
    permission = sum(1 for r in results if r["status"] == "PERMISSION_REQUIRED")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    if failed:
        overall = "FAIL"
    elif permission:
        overall = "PASS_WITH_PERMISSION_REQUIRED"
    elif missing:
        overall = "PASS_WITH_MISSING_FIXTURES"
    else:
        overall = "PASS"
    content = f"""# DotnetPOI {evidence_dir.name} macOS Office Evidence

- Project version: `{version}`
- Git revision: `{revision}`
- Captured: `{start.strftime('%Y-%m-%d %H:%M:%S %z')}` - `{end.strftime('%Y-%m-%d %H:%M:%S %z')}`
- Environment: macOS Microsoft Office apps
- Excel: `{versions.get('excel', 'unknown')}`
- Word: `{versions.get('word', 'unknown')}`
- PowerPoint: `{versions.get('powerpoint', 'unknown')}`
- Source root: `tools/manual-verification/generated-documents`
- Overall: `{overall}`
- Result counts: `{passing}` pass, `{missing}` missing fixture, `{permission}` permission required, `{failed}` fail

This evidence pass opens each available file in the corresponding macOS Microsoft Office app, treats open/reopen failures as failed cases, captures a screen PNG, and writes this index for GitHub review.

## Matrix

| kind | app | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---|---:|---:|---:|---:|---|---|
{table_rows(results)}

## Notes

- `MISSING` means generated manual documents are not present; run `dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj`.
- The original files are not modified; work copies are written under `{WORK_ROOT_NAME}/`.
- Password for generated encrypted files: `f`.
- Screenshots are captured with macOS `screencapture`; run permission/bootstrap mode first if preflight reports `PERMISSION_REQUIRED`.
"""
    (evidence_dir / INDEX_NAME).write_text(content, encoding="utf-8")
    return overall


def write_permission_required_index(evidence_dir, version, revision, preflight_result):
    content = f"""# DotnetPOI {evidence_dir.name} macOS Office Evidence

- Project version: `{version}`
- Git revision: `{revision}`
- Overall: `PERMISSION_REQUIRED`

macOS screen capture permission is required before evidence mode can run. This is intentionally not treated as a document compatibility failure.

Run:

```bash
tools/manual-verification/scripts/run-macos-office-permissions.sh
```

Then grant Screen Recording / Screen & System Audio Recording, Automation, and Accessibility permissions as prompted, and rerun evidence mode.

## Preflight

```text
returncode={preflight_result.returncode}
stdout={preflight_result.stdout}
stderr={preflight_result.stderr}
```
"""
    (evidence_dir / INDEX_NAME).write_text(content, encoding="utf-8")


def main():
    version = os.environ.get("DOTNETPOI_VERSION") or read_project_version()
    revision = os.environ.get("DOTNETPOI_REVISION") or "nogit"
    evidence_id = os.environ.get("DOTNETPOI_EVIDENCE_ID") or f"v{version}-{revision}-macos"
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
    screen_ok, screen_result = check_screencapture_permission(evidence_dir)
    if not screen_ok:
        write_permission_required_index(evidence_dir, version, revision, screen_result)
        log(session_log, "overall=PERMISSION_REQUIRED screen capture preflight failed")
        return 2

    versions = app_versions()
    log(session_log, f"version={version} revision={revision} office={versions}")

    results = []
    try:
        for ordinal, case in enumerate(case_matrix(), start=1):
            results.append(verify_case(case, ordinal, work_dir, image_dir, session_log))
    finally:
        close_all()

    end = now_jst()
    overall = write_index(evidence_dir, start, end, version, revision, versions, results)
    log(session_log, f"index written: {evidence_dir / INDEX_NAME}")
    log(session_log, f"overall={overall}")
    return 1 if any(r["status"] == "FAIL" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
