#!/usr/bin/env python3
import datetime as dt
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
OUT = ROOT / "tools" / "manual-verification" / "evidence" / "macos-permissions"
STATUS = OUT / "STATUS.md"


def now_jst():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))


def run(cmd, timeout=30):
    return subprocess.run(cmd, text=True, capture_output=True, timeout=timeout)


def check_screencapture():
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "screencapture-permission-test.png"
    if target.exists():
        target.unlink()
    result = run(["/usr/sbin/screencapture", "-x", str(target)], timeout=20)
    return result.returncode == 0 and target.exists(), result


def launch_office_apps():
    results = []
    for app in ("Microsoft Excel", "Microsoft Word", "Microsoft PowerPoint"):
        result = run(["/usr/bin/open", "-n", "-a", app], timeout=20)
        results.append((app, result))
    return results


def check_system_events():
    script = 'tell application "System Events" to get name of first process'
    return run(["osascript", "-e", script], timeout=20)


def check_key_events():
    # Escape is intentionally low impact. Sending any key is what exercises
    # Accessibility/Input Monitoring style permission gates.
    return run(
        [
            "osascript",
            "-e",
            'tell application "System Events"',
            "-e",
            "key code 53",
            "-e",
            "end tell",
        ],
        timeout=20,
    )


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    started = now_jst()
    office_results = launch_office_apps()
    screen_ok, screen_result = check_screencapture()
    system_events = check_system_events()
    key_events = check_key_events()

    lines = [
        "# macOS Office Permission Bootstrap",
        "",
        f"- Checked: `{started.strftime('%Y-%m-%d %H:%M:%S %z')}`",
        f"- Screen capture: `{'PASS' if screen_ok else 'PERMISSION_REQUIRED'}`",
        f"- System Events: `{'PASS' if system_events.returncode == 0 else 'PERMISSION_REQUIRED'}`",
        f"- Key events: `{'PASS' if key_events.returncode == 0 else 'PERMISSION_REQUIRED'}`",
        "",
        "## Office Launch",
        "",
        "| app | status | stderr |",
        "|---|---:|---|",
    ]
    for app, result in office_results:
        status = "PASS" if result.returncode == 0 else "CHECK_REQUIRED"
        stderr = (result.stderr or "").strip().replace("\n", "<br/>")
        lines.append(f"| {app} | {status} | {stderr} |")

    lines.extend(
        [
            "",
            "## Required Manual Grants",
            "",
            "If any item above is `PERMISSION_REQUIRED`, grant permissions in macOS System Settings and rerun this script.",
            "",
            "- Screen Recording / Screen & System Audio Recording: allow the app that runs Codex or the terminal host.",
            "- Accessibility: allow the app that runs Codex or the terminal host if password dialog automation is needed.",
            "- Automation: allow control of Microsoft Excel, Word, PowerPoint, and System Events when prompted.",
            "",
            "After permissions are granted, run:",
            "",
            "```bash",
            "tools/manual-verification/scripts/run-macos-office-evidence.sh",
            "```",
            "",
            "## Raw Results",
            "",
            "### screencapture",
            "",
            "```text",
            f"returncode={screen_result.returncode}",
            f"stdout={screen_result.stdout}",
            f"stderr={screen_result.stderr}",
            "```",
            "",
            "### System Events",
            "",
            "```text",
            f"returncode={system_events.returncode}",
            f"stdout={system_events.stdout}",
            f"stderr={system_events.stderr}",
            "```",
            "",
            "### Key Events",
            "",
            "```text",
            f"returncode={key_events.returncode}",
            f"stdout={key_events.stdout}",
            f"stderr={key_events.stderr}",
            "```",
            "",
        ]
    )
    STATUS.write_text("\n".join(lines), encoding="utf-8")
    print(f"Permission status written to: {STATUS}")
    return 0 if screen_ok and system_events.returncode == 0 and key_events.returncode == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
