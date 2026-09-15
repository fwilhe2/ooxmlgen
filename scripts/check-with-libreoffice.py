#!/usr/bin/env python3
"""Check the corpus against LibreOffice, which is the only oracle available for it.

Two claims, and the second is the one worth having:

  * every fixture whose manifest says `oracleOpens` converts, and every fixture that says it
    does not, does not. That field is a *measurement*, so this is a regression test on the
    corpus rather than a judgement about any importer;
  * nothing hangs, and nothing writes outside the output directory. The hostile family is the
    reason this script exists at all.

A mismatch does not mean the corpus is wrong. It means LibreOffice changed, which is worth
knowing before a downstream test suite discovers it.

usage: check-with-libreoffice.py <corpus-dir> [--timeout SECONDS]
"""

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# The entry names in hostile/zip-slip.xlsx traverse upwards. If a converter ever honours one,
# this is where it would land.
ESCAPE_MARKER = Path("/tmp/ooxmlgen-zip-slip-marker.xml")


def convert(soffice: str, profile: Path, out: Path, path: Path, timeout: int):
    start = time.monotonic()
    try:
        subprocess.run(
            [soffice, "--headless", f"-env:UserInstallation=file://{profile}",
             "--convert-to", "ods", "--outdir", str(out), str(path)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=timeout,
        )
        timed_out = False
    except subprocess.TimeoutExpired:
        timed_out = True
    elapsed = time.monotonic() - start
    produced = (out / (path.stem + ".ods")).is_file()
    return produced, timed_out, elapsed


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        print("no soffice on PATH — skipping (an oracle is not vendorable)")
        return 0

    manifest = json.loads((args.corpus / "manifest.json").read_text())
    ESCAPE_MARKER.unlink(missing_ok=True)

    problems: list[str] = []
    slow: list[tuple[str, float]] = []

    with tempfile.TemporaryDirectory() as scratch:
        profile = Path(scratch) / "profile"
        out = Path(scratch) / "out"
        out.mkdir()

        for fixture in manifest["fixtures"]:
            name = fixture["file"]
            expected = fixture.get("oracleOpens", True)
            produced, timed_out, elapsed = convert(
                soffice, profile, out, args.corpus / name, args.timeout)

            if timed_out:
                problems.append(f"{name}: timed out after {args.timeout}s")
            elif produced != expected:
                verb = "opened" if produced else "was refused"
                want = "open" if expected else "be refused"
                problems.append(f"{name}: {verb}; the manifest says it should {want}")

            if elapsed > 30:
                slow.append((name, elapsed))

    if ESCAPE_MARKER.exists():
        problems.append(f"{ESCAPE_MARKER}: a fixture wrote outside the output directory")
        ESCAPE_MARKER.unlink()

    print(f"{len(manifest['fixtures'])} fixtures converted with {Path(soffice).name}")
    for name, elapsed in slow:
        print(f"  slow  {name}: {elapsed:.0f}s")
    for problem in problems:
        print(f"  FAIL  {problem}")
    if problems:
        print(f"{len(problems)} problems")
        return 1
    print("every fixture behaved as the manifest says it does")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
