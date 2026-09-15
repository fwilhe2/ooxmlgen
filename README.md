# ooxmlgen

Generates a corpus of Excel workbooks for testing an importer against.

It exists for [grind](https://github.com/fwilhe2/grind)'s `.xlsx` import filter, whose plan
(`doc/xlsx-import.md`) asks for three kinds of fixture that a corpus of somebody else's bug
reports will not contain:

> Hand-built fixtures for the boundaries that no corpus reliably contains … Fixtures for the
> real-world section, which is the half a LibreOffice corpus is least likely to cover, since
> its files are minimal bug reproductions rather than things Excel wrote … A Strict fixture,
> one file, asserting only that it imports.

That is what this writes. Nothing here is specific to grind, though: the manifest describes
what each file contains and what a conversion of it should produce, in ODF's vocabulary, and
any importer can be measured against it.

## Running it

```
dotnet run -- --out corpus
```

```
  -o, --out <dir>     where to write the corpus (default: ./corpus)
      --only <a,b>    generate only these families
      --skip <a,b>    generate everything but these families
      --no-clean      keep whatever is already in the output directory
      --list          list the families and what each covers
```

The output is one directory per family, a `manifest.json`, and a `README.md` indexing the lot.
Nothing is committed: **this program is the corpus's source**, the files are never hand-edited,
and `corpus/` is in `.gitignore`. Vendor the output into whatever consumes it.

## What it generates

76 fixtures, about 5 MiB, in nine families. Each file exercises one construct and is named
for it, so a failure names its own cause.

| Family | | |
| --- | --- | --- |
| `values/` | X1 | cell types, numeric literals, strings and Unicode, rich text, both date systems and the 1900 leap-year bug, times |
| `formulas/` | X2 | references, operators and precedence, functions, shared-formula groups, `_xlfn.`, the excluded classes, divergent semantics |
| `numfmt/` | X3 | every built-in id, section rules, custom codes, currency and locale tags, date/time pieces, conditions |
| `styles/` | X4 | fonts, the four ways to spell a colour, fills and patterns, all thirteen borders, alignment, named styles |
| `geometry/` | X4 | column widths, row heights, hidden and outlined tracks, frozen and split panes |
| `document/` | X5 | defined names, sheet order and visibility, merges, filters, tables, comments, conditional formats, validation, protection, `.xlsm`, external links, charts, pivot tables |
| `realworld/` | X0 | implicit `r`, `mc:AlternateContent`, Strict and mixed namespaces, `xml:space`, lying counts, awkward part targets, BOMs, odd packaging |
| `hostile/` | X0 | zip bomb, ten thousand parts, billion laughs, XXE, zip slip, CFB-encrypted, truncated, deeply nested |
| `scale/` | X1 | half a million cells, and six cells at the far corners of a sheet |

## The manifest

Every cell is stated twice: what the file literally contains, and what a conversion of it
should produce. Both halves are written at the point the cell is generated, so the manifest
cannot drift away from the bytes.

```json
{
  "file": "values/dates-1900.xlsx",
  "family": "values",
  "milestone": "X1",
  "covers": "the 1900 epoch and the phantom 1900-02-29",
  "flavour": "Transitional",
  "oracleOpens": true,
  "expectDropped": {},
  "sheets": [{
    "name": "Dates1900",
    "cells": [{
      "ref": "B2", "raw": "1", "style": 1,
      "kind": "Date", "value": "1900-01-01", "display": "1900-01-01",
      "oracle": "1899-12-31",
      "note": "the epoch, as Excel displays it"
    }]
  }]
}
```

`expectDropped`'s keys and `error`'s value are grind's `Dropped` and `Error` variants spelled
exactly as the Rust does, so a consumer matches on them without a translation table.

### `oracle`

Where a cell carries an `oracle`, LibreOffice's own conversion of that cell produces something
other than `value` — measured, not guessed. A test comparing an importer against
`soffice --convert-to ods` will disagree on exactly those cells, through no fault of its own.
Five such divergences turned up while building this:

| | grind's plan / Excel's meaning | LibreOffice |
| --- | --- | --- |
| 1900 serials 1–60 | corrected by +1: serial 1 is `1900-01-01` | no correction at all: serial 1 is `1899-12-31` |
| cross-sheet reference | `[Data.A1]` | `[$Data.A1]` |
| `_xlfn.XLOOKUP(…)` | prefix stripped → `XLOOKUP(…)`, then reported unknown | `COM.MICROSOFT.XLOOKUP(…)` |
| shared formula shifted off the sheet | — | `[#REF!]*10`: only the reference breaks, not the expression |
| `hh:mm:ss` vs `[h]:mm:ss` | — | identical duration; only the *display* differs |

The last two are corrections to this corpus rather than to grind — the first drafts of both
fixtures asserted the wrong thing, and converting them with the real oracle is what caught it.

## Checking it

```
python3 scripts/check-corpus.py corpus           # the files match the manifest
python3 scripts/check-with-libreoffice.py corpus # and LibreOffice agrees about which open
```

The first checks hashes, sizes, orphans, and the structural claims — is it a zip, does it carry
a CFB signature, are the duplicate entries the declared ones. The second converts all 76 with
LibreOffice and asserts each behaves as `oracleOpens` says, that nothing hangs, and that the
zip-slip fixture writes nothing outside its output directory. It skips cleanly when `soffice`
is not on `PATH`, because an oracle is not vendorable. Both run in CI.

`oracleOpens` is a measurement, not a judgement: several files LibreOffice opens a stricter
reader may rightly refuse, and `realworld/duplicate-entry.xlsx` is one it refuses that a better
reader could handle.

## A note on the hostile family

These are shapes a parser has to decline, not working exploits: each is small, inert, and
paired with the verdict it should produce. The XXE fixture points at a `.invalid` host, which
by RFC 2606 can never resolve. The zip-slip entry names target `/tmp` with a marker filename
so that a reader which does follow them fails visibly in testing rather than quietly in
production.

None of them should be an out-of-memory, a hang, a stack overflow, a network request, or a
write outside the output directory. Measured against LibreOffice, five of the twelve are opened
rather than refused — including the billion-laughs and XXE files, which it reads without
expanding or fetching anything. Refusing is acceptable; surviving is the requirement.
