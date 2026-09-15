#!/usr/bin/env python3
"""Check a generated corpus against its own manifest.

The generator is the only thing that writes these files, so "it still runs" is a weak claim.
This checks that what it wrote is what it said it wrote: every fixture present, every hash
right, nothing on disk that the manifest does not know about, and the structural claims —
is it a zip, does it carry a CFB signature, are there duplicate entries — matching the
`error` field beside them.

usage: check-corpus.py <corpus-dir>
"""

import hashlib
import json
import sys
import zipfile
from pathlib import Path

CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# Fixtures that are a valid archive whose *contents* are the problem. For these, "opens as a
# zip" is expected and is not evidence against the declared error.
VALID_ZIP_WITH_BAD_CONTENTS = {
    "hostile/zip-slip.xlsx",
    "hostile/no-workbook-part.xlsx",
    "hostile/empty-zip.xlsx",
}


def main(root: Path) -> int:
    manifest = json.loads((root / "manifest.json").read_text())
    fixtures = manifest["fixtures"]
    problems: list[str] = []

    def fail(fixture: str, why: str) -> None:
        problems.append(f"{fixture}: {why}")

    listed = set()
    for fixture in fixtures:
        name = fixture["file"]
        listed.add(name)
        path = root / name

        if not path.is_file():
            fail(name, "listed in the manifest and not on disk")
            continue

        body = path.read_bytes()
        if len(body) != fixture["bytes"]:
            fail(name, f"size is {len(body)}, manifest says {fixture['bytes']}")
        digest = hashlib.sha256(body).hexdigest()
        if digest != fixture["sha256"]:
            fail(name, f"sha256 is {digest}, manifest says {fixture['sha256']}")

        error = fixture.get("error")

        if error == "Encrypted":
            if body[:8] != CFB_SIGNATURE:
                fail(name, "declares Error::Encrypted without a CFB signature")
            continue

        try:
            entries = zipfile.ZipFile(path).namelist()
            is_zip = True
        except (zipfile.BadZipFile, OSError):
            entries = []
            is_zip = False

        if error == "Package":
            # These are meant to be unreadable as packages. Some are not archives at all;
            # others are valid archives whose contents are the problem.
            if is_zip and name not in VALID_ZIP_WITH_BAD_CONTENTS:
                fail(name, "declares Error::Package but is a readable zip")
        elif not is_zip:
            fail(name, "should be a readable zip and is not")

        duplicates = len(entries) != len(set(entries))
        declared = fixture.get("duplicateEntries", False)
        if duplicates and not declared:
            fail(name, "has duplicate zip entries and does not declare them")
        if declared and not duplicates:
            fail(name, "declares duplicate entries that are not there")

    on_disk = {
        str(p.relative_to(root)).replace("\\", "/")
        for p in root.rglob("*")
        if p.is_file() and p.suffix in {".xlsx", ".xlsm"}
    }
    for orphan in sorted(on_disk - listed):
        fail(orphan, "on disk and not in the manifest")

    cells = sum(len(s["cells"]) for f in fixtures for s in f.get("sheets", []))
    print(f"{len(fixtures)} fixtures, {cells} asserted cells")
    for problem in problems:
        print(f"  FAIL {problem}")
    if problems:
        print(f"{len(problems)} problems")
        return 1
    print("corpus is consistent with its manifest")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__, file=sys.stderr)
        raise SystemExit(2)
    raise SystemExit(main(Path(sys.argv[1])))
