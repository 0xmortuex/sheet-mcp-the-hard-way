"""Run every module's main.py in a fresh process and require exit code 0 plus the MODULE OK line.

Each main.py checks its own claims with assert, so a chapter whose code drifts out of sync with
its README fails here. Children run with PYTHONIOENCODING=utf-8: Windows CI consoles default to
cp1252 and would crash on the Turkish text these modules print.

Usage: python verify.py            # all modules
       python verify.py 04 06      # just these
"""

from __future__ import annotations

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))


def run(module: str) -> tuple[bool, str, float]:
    script = os.path.join(ROOT, module, "main.py")
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    start = time.monotonic()
    try:
        proc = subprocess.run(
            [sys.executable, script], capture_output=True, text=True, encoding="utf-8",
            timeout=180, cwd=os.path.join(ROOT, module), env=env,
        )
    except subprocess.TimeoutExpired:
        return False, "timed out after 180s", time.monotonic() - start
    output = proc.stdout + proc.stderr
    ok = proc.returncode == 0 and "MODULE OK" in proc.stdout
    return ok, output, time.monotonic() - start


def main() -> int:
    wanted = sys.argv[1:]
    modules = sorted(
        d for d in os.listdir(ROOT)
        if os.path.isfile(os.path.join(ROOT, d, "main.py")) and (not wanted or any(d.startswith(w) for w in wanted))
    )
    if not modules:
        print("no modules matched")
        return 1
    failed = []
    for m in modules:
        ok, output, secs = run(m)
        print(f"{'PASS' if ok else 'FAIL'}  {m}  ({secs:.1f}s)")
        if not ok:
            failed.append(m)
            print("    " + output.strip().replace("\n", "\n    "))
    print(f"\n{len(modules) - len(failed)}/{len(modules)} modules verified")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
