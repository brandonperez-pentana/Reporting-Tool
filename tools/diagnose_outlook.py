"""Work out why Outlook automation fails on this machine.

`Invalid class string` (-2147221005) means Windows could not resolve the ProgID
"Outlook.Application". Classic Outlook being installed and open does not guarantee that:
the ProgID lives in a registry view that the running process has to be able to see, and
there are several ordinary ways for it not to.

This reports what is actually true here rather than guessing. It only reads the registry
and, at the end, tries to create the Outlook object. It sends nothing, opens nothing and
changes nothing.

    python tools/diagnose_outlook.py
"""

from __future__ import annotations

import os
import platform
import struct
import sys
from pathlib import Path

PROGIDS = ("Outlook.Application", "Outlook.Application.16", "Outlook.Application.15")


def heading(text: str) -> None:
    print(f"\n{text}\n{'-' * len(text)}")


def pe_bitness(path: Path) -> str:
    """Read the PE header to tell a 32-bit binary from a 64-bit one."""
    try:
        with path.open("rb") as handle:
            if handle.read(2) != b"MZ":
                return "not a Windows binary"
            handle.seek(0x3C)
            pe_offset = struct.unpack("<I", handle.read(4))[0]
            handle.seek(pe_offset)
            if handle.read(4) != b"PE\0\0":
                return "unreadable PE header"
            machine = struct.unpack("<H", handle.read(2))[0]
    except OSError as exc:
        return f"unreadable ({exc})"
    return {0x014C: "32-bit", 0x8664: "64-bit", 0xAA64: "ARM64"}.get(machine, f"unknown (0x{machine:04X})")


def report_process() -> None:
    heading("This Python process")
    bits = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    print(f"  Python           {platform.python_version()}  {bits}")
    print(f"  Executable       {sys.executable}")
    print(f"  Virtual env      {'yes' if sys.prefix != sys.base_prefix else 'no'}")
    if os.name == "nt":
        try:
            import ctypes

            elevated = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:  # noqa: BLE001
            elevated = False
        print(f"  Running as admin {'YES' if elevated else 'no'}")
        if elevated:
            print("    ^ Office registers its COM classes per user. An elevated process can")
            print("      read a different HKEY_CURRENT_USER, which hides the ProgID. Try again")
            print("      from a normal, non-elevated PowerShell window.")
    try:
        import win32com  # noqa: F401

        print("  pywin32          installed")
    except ImportError:
        print("  pywin32          NOT INSTALLED  ->  pip install pywin32")


def registry_lookup() -> list[str]:
    """Where each ProgID resolves, in every registry view that matters."""
    heading("Registry: is Outlook.Application registered?")
    if os.name != "nt":
        print("  (not Windows — nothing to read)")
        return []
    import winreg

    found: list[str] = []
    roots = [
        ("HKEY_CLASSES_ROOT", winreg.HKEY_CLASSES_ROOT, 0),
        ("HKEY_CLASSES_ROOT (32-bit view)", winreg.HKEY_CLASSES_ROOT, winreg.KEY_WOW64_32KEY),
        ("HKEY_CLASSES_ROOT (64-bit view)", winreg.HKEY_CLASSES_ROOT, winreg.KEY_WOW64_64KEY),
        ("HKEY_CURRENT_USER\\Software\\Classes", winreg.HKEY_CURRENT_USER, 0),
        ("HKEY_LOCAL_MACHINE\\Software\\Classes", winreg.HKEY_LOCAL_MACHINE, 0),
    ]
    for progid in PROGIDS:
        print(f"\n  {progid}")
        for label, hive, flag in roots:
            prefix = "Software\\Classes\\" if "Software" in label else ""
            try:
                with winreg.OpenKey(hive, f"{prefix}{progid}\\CLSID", 0, winreg.KEY_READ | flag) as key:
                    clsid = winreg.QueryValue(key, "")
                print(f"    found in {label}  ->  {clsid}")
                found.append(clsid)
            except OSError:
                print(f"    absent   in {label}")
    return found


def server_for(clsid: str) -> None:
    """Which executable a CLSID points at, and whether it is really there."""
    import winreg

    for label, flag in (("64-bit view", winreg.KEY_WOW64_64KEY), ("32-bit view", winreg.KEY_WOW64_32KEY)):
        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", 0, winreg.KEY_READ | flag
            ) as key:
                command = winreg.QueryValue(key, "")
        except OSError:
            print(f"    {label}: no LocalServer32")
            continue
        executable = command.strip('"').split('" ')[0].strip('"').split(" /")[0]
        path = Path(executable)
        exists = path.is_file()
        print(f"    {label}: {command}")
        print(f"      exists: {'yes' if exists else 'NO — this is the problem'}"
              f"{'  |  ' + pe_bitness(path) if exists else ''}")


def try_dispatch() -> list[tuple[str, bool]]:
    """Try each way of getting the Application object; report and return the outcomes."""
    heading("Can this process actually start Outlook?")
    results: list[tuple[str, bool]] = []
    if os.name != "nt":
        print("  (not Windows — skipped)")
        return results
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("  pywin32 is not installed, so no.")
        return results

    pythoncom.CoInitialize()
    try:
        for name, call in (
            ("Dispatch('Outlook.Application')", lambda: win32com.client.Dispatch("Outlook.Application")),
            ("DispatchEx('Outlook.Application')", lambda: win32com.client.DispatchEx("Outlook.Application")),
            ("GetActiveObject('Outlook.Application')", lambda: win32com.client.GetActiveObject("Outlook.Application")),
        ):
            try:
                call()
            except Exception as exc:  # noqa: BLE001
                print(f"  {name}\n      FAILED: {exc}")
                results.append((name, False))
            else:
                print(f"  {name}\n      OK — Reporter can use this.")
                results.append((name, True))
    finally:
        pythoncom.CoUninitialize()
    return results


def find_outlook() -> None:
    heading("Classic Outlook on disk")
    if os.name != "nt":
        print("  (not Windows — skipped)")
        return
    roots = [Path(os.environ.get(var, "")) for var in ("ProgramFiles", "ProgramFiles(x86)") if os.environ.get(var)]
    hits = []
    for root in roots:
        for pattern in ("Microsoft Office/root/Office*/OUTLOOK.EXE", "Microsoft Office/Office*/OUTLOOK.EXE"):
            hits.extend(sorted(root.glob(pattern)))
    if not hits:
        print("  No OUTLOOK.EXE found under Program Files. Classic Outlook may not be installed.")
        return
    for path in hits:
        print(f"  {path}\n      {pe_bitness(path)}")


def verdict(found: list[str], dispatch: list[tuple[str, bool]]) -> None:
    """Read the actual results, not just the registry — a working Dispatch settles it."""
    heading("What to do")
    bits = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    worked = [name for name, ok in dispatch if ok]

    if worked:
        print(f"  Outlook automation works from this Python: {len(worked)} of "
              f"{len(dispatch)} calls above succeeded.")
        print("  Nothing on this machine needs changing. If Reporter still shows the")
        print("  COM error:")
        print("   1. In the browser app, go to stage 2 and click Re-check. The Outlook")
        print("      probe runs once per page load, so an earlier failure stays on")
        print("      screen until you re-check or reload.")
        print("   2. Start Reporter from this same venv, and from the same kind of")
        print("      window (elevated or not) you ran this from.")
        print("   3. Leave classic Outlook open while sending. A failure that clears")
        print("      once Outlook is running usually means it was not running before.")
        return

    if not found:
        print("  The ProgID is not registered anywhere this process can see it.")
        print("  In order of likelihood:")
        print("   1. Run this from a NON-elevated PowerShell if you used an admin one.")
        print("   2. Office repair: Settings > Apps > Microsoft 365 > Modify > Quick Repair.")
        print("      That re-registers the COM classes; /regserver no longer exists in")
        print("      Click-to-Run Office, which is why it reported an invalid switch.")
        print("   3. Activate Office if it is prompting — an unlicensed install can end up")
        print("      partially registered.")
        return

    if dispatch:
        print(f"  The ProgID resolves and this is a {bits} Python, but every call above")
        print("  failed. The usual cause is a bitness mismatch between Python and Office")
    else:
        print(f"  The ProgID resolves, and this is a {bits} Python.")
        print("  If Dispatch failed when Reporter ran, the usual cause is a bitness mismatch")
    print("  — compare the two lines under 'Classic Outlook on disk' with this process.")
    print("  Matching them fixes it: install the Python that matches Office, and recreate")
    print("  the venv with it.")


def main() -> int:
    print("Reporter — Outlook automation diagnostics")
    report_process()
    found = registry_lookup()
    if found:
        heading("Where that CLSID points")
        for clsid in dict.fromkeys(found):
            print(f"  {clsid}")
            server_for(clsid)
    find_outlook()
    dispatch = try_dispatch()
    verdict(found, dispatch)
    print("\nPaste this whole output back if you want help reading it.\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
