using OoxmlGen.Raw;

namespace OoxmlGen.Families;

/// <summary>
/// X0 — what real producers write that the spec permits and nobody expects.
///
/// This is the family a corpus of bug reproductions is least likely to cover, because those
/// files are minimal reductions rather than things a program in the wild actually wrote.
/// Every fixture here is a one-line rule in a reader and a bug report each if the rule is
/// missing. They are built from XML text rather than through an object model, because a
/// well-behaved object model will not emit any of them.
/// </summary>
public static class RealWorld
{
    public static void Generate(Corpus corpus)
    {
        ImplicitReferences(corpus);
        AlternateContent(corpus);
        IgnorableAndMustUnderstand(corpus);
        StrictNamespaces(corpus);
        MixedNamespaces(corpus);
        PreservedWhitespace(corpus);
        LyingCounts(corpus);
        PartTargets(corpus);
        BomAndDeclaration(corpus);
        UnusualPackaging(corpus);
        SheetOrder(corpus);
        DefaultNamespace(corpus);
    }

    static void ImplicitReferences(Corpus corpus)
    {
        var spec = New("realworld/implicit-refs.xlsx", "X0", "rows and cells with no r attribute");
        spec.Note("Position is implicit when `r` is absent: the next row down, the next column " +
                  "across. Excel always writes `r`; Apache POI and SheetJS do not always. A reader " +
                  "that unwraps `r` rather than defaulting it panics on a file a large part of the " +
                  "ecosystem produces.");
        spec.Note("Row 3 is the case that separates a real implementation from a guess: it mixes " +
                  "the two, so the implicit column counter has to resume from the last explicit " +
                  "`r` rather than from the count of cells seen.");
        spec.Note("Row 5 has an explicit `r` that skips to row 5 after two implicit rows, proving " +
                  "the row counter is a position and not an increment.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook { SharedStrings = null };
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1">
                    <c r="A1"><v>11</v></c>
                    <c r="B1"><v>12</v></c>
                    <c r="C1"><v>13</v></c>
                  </row>
                  <row>
                    <c><v>21</v></c>
                    <c><v>22</v></c>
                    <c><v>23</v></c>
                  </row>
                  <row>
                    <c><v>31</v></c>
                    <c r="D3"><v>34</v></c>
                    <c><v>35</v></c>
                  </row>
                  <row r="5">
                    <c r="A5"><v>51</v></c>
                    <c><v>52</v></c>
                  </row>
                </sheetData>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", "11", null), ("B1", "12", null), ("C1", "13", null),
            ("A2", "21", "the row has no `r`: it is the one after row 1"),
            ("B2", "22", null), ("C2", "23", null),
            ("A3", "31", "implicit again, and the row is 3"),
            ("D3", "34", "an explicit `r` mid-row jumps the column counter to D"),
            ("E3", "35", "and the next implicit cell resumes from D, not from the third cell"),
            ("A5", "51", "an explicit row 5 after implicit rows 2 and 3 — row 4 does not exist"),
            ("B5", "52", null));
    }

    static void AlternateContent(Corpus corpus)
    {
        var spec = New("realworld/mce-alternate-content.xlsx", "X0",
                       "mc:AlternateContent — the Choice must be skipped and the Fallback read");
        spec.Note("The rule: this reader `Requires` nothing, so every `mc:Choice` is skipped and " +
                  "the `mc:Fallback`, if there is one, is read in its place. There are three ways " +
                  "to get this wrong and all three are silent — take the Choice, take neither, or " +
                  "take both and double whatever was inside.");
        spec.Note("The Choice and the Fallback here hold *different values on purpose*. A reader " +
                  "that takes the wrong branch produces 999 instead of 1, which is visible; a " +
                  "fixture where both branches agreed would pass either way.");
        spec.Note("Row 3's AlternateContent has no Fallback at all, so the whole element " +
                  "contributes nothing. Row 4 nests one inside another.");
        spec.Note("`mc:AlternateContent` can wrap almost anything — a sparkline group, a slicer, " +
                  "a data validation, a `sheetPr` — so the rule belongs in the element dispatcher " +
                  "once, not as a match arm at each site that might contain one.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.ExtraNamespaces.Add($"""xmlns:mc="{RawWorkbook.MarkupCompatibility}" """.TrimEnd());
            book.ExtraNamespaces.Add("""xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" """.TrimEnd());
            book.RootAttributes = """ mc:Ignorable="x14" """.TrimEnd();
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1">
                    <c r="A1" t="inlineStr"><is><t>choice vs fallback</t></is></c>
                  </row>
                  <mc:AlternateContent>
                    <mc:Choice Requires="x14">
                      <row r="2"><c r="A2"><v>999</v></c></row>
                    </mc:Choice>
                    <mc:Fallback>
                      <row r="2"><c r="A2"><v>1</v></c></row>
                    </mc:Fallback>
                  </mc:AlternateContent>
                  <mc:AlternateContent>
                    <mc:Choice Requires="x14">
                      <row r="3"><c r="A3"><v>998</v></c></row>
                    </mc:Choice>
                  </mc:AlternateContent>
                  <mc:AlternateContent>
                    <mc:Choice Requires="x14">
                      <row r="4"><c r="A4"><v>997</v></c></row>
                    </mc:Choice>
                    <mc:Fallback>
                      <mc:AlternateContent>
                        <mc:Choice Requires="x14">
                          <row r="4"><c r="A4"><v>996</v></c></row>
                        </mc:Choice>
                        <mc:Fallback>
                          <row r="4"><c r="A4"><v>4</v></c></row>
                        </mc:Fallback>
                      </mc:AlternateContent>
                    </mc:Fallback>
                  </mc:AlternateContent>
                  <row r="5"><c r="A5"><v>5</v></c></row>
                </sheetData>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", null, "an ordinary cell before any of it"),
            ("A2", "1", "the Fallback's value. 999 means the Choice was taken; two cells at A2 " +
                        "means both were"),
            ("A3", null, "no Fallback, so the whole AlternateContent contributes nothing and A3 " +
                         "does not exist"),
            ("A4", "4", "a nested AlternateContent: the outer Fallback holds another one, whose " +
                        "own Fallback is the value that survives"),
            ("A5", "5", "an ordinary cell after all of it, proving the dispatcher resumed"));
        spec.Sheets[0].Cells[0].Kind = ValueKind.Text;
        spec.Sheets[0].Cells[0].Value = "choice vs fallback";
        spec.Sheets[0].Cells[2].Kind = ValueKind.Empty;
    }

    static void IgnorableAndMustUnderstand(Corpus corpus)
    {
        var spec = New("realworld/mce-ignorable.xlsx", "X0", "mc:Ignorable prefixes and mc:MustUnderstand");
        spec.ExpectMustUnderstand.Add("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        spec.Note("`mc:Ignorable` lists prefixes whose *attributes* a consumer may skip. Nothing " +
                  "needs building for it: an attribute in a namespace no table knows already " +
                  "misses every lookup. It is here so that nobody builds something.");
        spec.Note("`mc:MustUnderstand` is a producer asserting a consumer cannot proceed without a " +
                  "namespace. It is recorded by name and is *not* a refusal — in a spreadsheet it " +
                  "guards a feature, and cell values are what an import is for. Refusing a whole " +
                  "workbook because one slicer needs x14 is the same mistake as refusing it over " +
                  "one chart.");
        spec.Note("Even LibreOffice's own xlsx output declares seven namespaces beyond the one it " +
                  "uses. The median real document is not kinder, which is the argument for a " +
                  "tolerant reader rather than a schema-shaped one.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.ExtraNamespaces.AddRange(
            [
                $"""xmlns:mc="{RawWorkbook.MarkupCompatibility}" """.TrimEnd(),
                """xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" """.TrimEnd(),
                """xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" """.TrimEnd(),
                """xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" """.TrimEnd(),
                """xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" """.TrimEnd(),
                """xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" """.TrimEnd(),
                """xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" """.TrimEnd(),
                """xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" """.TrimEnd(),
            ]);
            book.RootAttributes =
                """ mc:Ignorable="x14ac xr xr2 xr3 xr6 xr10" mc:MustUnderstand="x15" xr:uid="{00000000-0001-0000-0000-000000000000}" """
                .TrimEnd();
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1" x14ac:dyDescent="0.25" xr:uid="{11111111-2222-3333-4444-555555555555}">
                    <c r="A1"><v>1</v></c>
                    <c r="B1" xr3:uid="unknown-attribute"><v>2</v></c>
                  </row>
                  <row r="2" spans="1:2" x14ac:dyDescent="0.25">
                    <c r="A2" t="inlineStr"><is><t>attributes in unknown namespaces are skipped</t></is></c>
                  </row>
                </sheetData>
                <x15:someExtensionElement/>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", "1", "the row carries x14ac and xr attributes that no table knows"),
            ("B1", "2", "an attribute in a namespace declared but not in the Ignorable list — " +
                        "still skipped, because skipping is the default rather than a permission"),
            ("A2", null, "an element in the x15 namespace follows sheetData and is skipped whole"));
        spec.Sheets[0].Cells[2].Kind = ValueKind.Text;
        spec.Sheets[0].Cells[2].Value = "attributes in unknown namespaces are skipped";
    }

    static void StrictNamespaces(Corpus corpus)
    {
        var spec = New("realworld/strict.xlsx", "X0", "ISO Strict namespaces throughout");
        spec.Flavour = Flavour.Strict;
        spec.Note("Strict is an opt-in save format almost nothing produces: Excel has written " +
                  "Transitional by default since 2007 and still does, and so does every other " +
                  "generator in the wild. This file is not a supported configuration — it is a " +
                  "proof that the namespace table is a table.");
        spec.Note("The whole difference that reaches a reader is which URIs it recognises. The " +
                  "package relationships, the content types and markup compatibility are Parts 2 " +
                  "and 3 and do not vary, which is why only three rows of the table change.");
        spec.Note("Recognising Strict costs about ten lines, and is worth them not because anyone " +
                  "sends Strict files but because writing the table forces the reader to be " +
                  "namespace-driven — which is the property that makes the *next* namespace inert " +
                  "for free.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook { UseStrict = true };
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1">
                    <c r="A1" t="inlineStr"><is><t>strict</t></is></c>
                    <c r="B1"><v>42</v></c>
                  </row>
                  <row r="2">
                    <c r="A2" s="1"><v>45368</v></c>
                    <c r="B2"><f>B1*2</f><v>84</v></c>
                  </row>
                </sheetData>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", null, "an inline string under the purl.oclc.org namespace"),
            ("B1", "42", null),
            ("A2", null, "a date, so the styles part has to resolve under Strict as well"),
            ("B2", "84", "and a formula, to prove nothing else in the reader is namespace-bound"));
        spec.Sheets[0].Cells[0].Kind = ValueKind.Text;
        spec.Sheets[0].Cells[0].Value = "strict";
        spec.Sheets[0].Cells[2].Kind = ValueKind.Date;
        spec.Sheets[0].Cells[2].Value = "2024-03-17";
        spec.Sheets[0].Cells[3].Excel = "B1*2";
        spec.Sheets[0].Cells[3].Formula = "[.B1]*2";
    }

    static void MixedNamespaces(Corpus corpus)
    {
        var spec = New("realworld/mixed-flavour.xlsx", "X0", "Transitional workbook, Strict worksheet");
        spec.Flavour = Flavour.Mixed;
        spec.Note("Files that mix the two families exist. Resolving per element rather than " +
                  "per document handles them for free; branching on a flavour decided once at the " +
                  "top of the file does not, and refusing them is the worst of the three.");
        spec.Note("`Flavour` is a fact the report states, not a mode the reader runs in. This is " +
                  "the file that proves the difference matters.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            // The workbook part is Transitional; the worksheet below overrides its own root.
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1"><v>1</v></c></row>
                </sheetData>
                """);
            book.ExtraParts.Add(("xl/worksheets/sheet2.xml", $"""
                <worksheet xmlns="{RawWorkbook.Strict}" xmlns:r="{RawWorkbook.StrictRelationships}">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>this sheet is Strict</t></is></c>
                      <c r="B1"><v>2</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """));
            book.ExtraOverrides.Add(("/xl/worksheets/sheet2.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            book.SheetPaths.Add("xl/worksheets/sheet2.xml");
            book.SheetNames.Add("Sheet2");
            book.SheetBodies.Add("<sheetData/>");
            book.Save(path);
        });

        SheetOf(spec, "Sheet1", ("A1", "1", "the Transitional half"));

        var strict = new SheetSpec { Name = "Sheet2" };
        strict.Cells.Add(new CellSpec
        {
            Ref = "A1", T = "inlineStr", Kind = ValueKind.Text, Value = "this sheet is Strict",
            Note = "the same package, the same workbook part, a worksheet under the other namespace",
        });
        strict.Cells.Add(new CellSpec { Ref = "B1", Raw = "2", Kind = ValueKind.Number, Value = "2" });
        spec.Sheets.Add(strict);
    }

    static void PreservedWhitespace(Corpus corpus)
    {
        var spec = New("realworld/xml-space-preserve.xlsx", "X0", "xml:space on shared and inline strings");
        spec.Note("Leading and trailing spaces in a string are meaningful and are lost unless the " +
                  "element says `xml:space=\"preserve\"`. A reader that trims corrupts text " +
                  "silently, which is the worst failure mode a text importer has.");
        spec.Note("Row 4 is the trap in the other direction: a string *without* `xml:space` whose " +
                  "content has surrounding whitespace. Strictly it may be collapsed, and real " +
                  "producers write it both ways, so the safe reading is to preserve what is there " +
                  "and never to trim on the reader's own initiative.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook
            {
                SharedStrings = $"""
                    <sst xmlns="{RawWorkbook.Transitional}" count="5" uniqueCount="5">
                      <si><t xml:space="preserve">  two leading spaces</t></si>
                      <si><t xml:space="preserve">two trailing spaces  </t></si>
                      <si><t xml:space="preserve">   </t></si>
                      <si><t>  no preserve attribute  </t></si>
                      <si><t xml:space="preserve">a	tab and a
                    newline</t></si>
                    </sst>
                    """,
            };
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1" t="s"><v>0</v></c></row>
                  <row r="2"><c r="A2" t="s"><v>1</v></c></row>
                  <row r="3"><c r="A3" t="s"><v>2</v></c></row>
                  <row r="4"><c r="A4" t="s"><v>3</v></c></row>
                  <row r="5"><c r="A5" t="s"><v>4</v></c></row>
                  <row r="6"><c r="A6" t="inlineStr"><is><t xml:space="preserve">  inline, preserved  </t></is></c></row>
                </sheetData>
                """);
            book.Save(path);
        });

        var sheet = new SheetSpec { Name = "Sheet1" };
        void Text(string reference, string value, string note) => sheet.Cells.Add(new CellSpec
        {
            Ref = reference,
            T = "s",
            Kind = ValueKind.Text,
            Value = value,
            Note = note,
        });

        Text("A1", "  two leading spaces", "both spaces survive");
        Text("A2", "two trailing spaces  ", "both spaces survive");
        Text("A3", "   ", "a string that is nothing but whitespace — not an empty cell");
        Text("A4", "  no preserve attribute  ", "no xml:space, and still not the reader's to trim");
        Text("A5", "a\ttab and a\n                    newline", "a tab and a literal newline, with the " +
             "source indentation inside them — whitespace in XML text is content");
        Text("A6", "  inline, preserved  ", "inline text needs the attribute exactly as a shared string does");
        spec.Sheets.Add(sheet);
    }

    static void LyingCounts(Corpus corpus)
    {
        var spec = New("realworld/lying-counts.xlsx", "X0", "count, uniqueCount and dimension as claims rather than facts");
        spec.Note("`count` and `uniqueCount` on `sst`, and `dimension` on a sheet, are claims. " +
                  "Never preallocate from them and never trust `dimension` as a bound: " +
                  "`A1:XFD1048576` is a claim a great many generators make, and materialising it " +
                  "is 17 billion cells for a file with four.");
        spec.Note("Here the string counts are wrong in both directions at once — `count` too high " +
                  "and `uniqueCount` too low — so a reader that sizes a table from either produces " +
                  "the wrong answer, and one that sizes from the elements it actually reads does not.");
        spec.Note("`spans` on a row is the same kind of claim at a smaller scale, and is wrong here too.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook
            {
                SharedStrings = $"""
                    <sst xmlns="{RawWorkbook.Transitional}" count="9999" uniqueCount="1">
                      <si><t>first</t></si>
                      <si><t>second</t></si>
                      <si><t>third</t></si>
                    </sst>
                    """,
            };
            book.SheetBodies.Add("""
                <dimension ref="A1:XFD1048576"/>
                <sheetData>
                  <row r="1" spans="1:16384">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1" t="s"><v>1</v></c>
                  </row>
                  <row r="2" spans="5:9">
                    <c r="A2" t="s"><v>2</v></c>
                    <c r="B2"><v>42</v></c>
                  </row>
                </sheetData>
                """);
            book.Save(path);
        });

        var sheet = new SheetSpec { Name = "Sheet1" };
        sheet.Cells.Add(new CellSpec
        {
            Ref = "A1", T = "s", Kind = ValueKind.Text, Value = "first",
            Note = "dimension claims the whole sheet; four cells exist",
        });
        sheet.Cells.Add(new CellSpec { Ref = "B1", T = "s", Kind = ValueKind.Text, Value = "second" });
        sheet.Cells.Add(new CellSpec
        {
            Ref = "A2", T = "s", Kind = ValueKind.Text, Value = "third",
            Note = "uniqueCount says 1 and there are 3 — a reader that trusts it reads 'first' here",
        });
        sheet.Cells.Add(new CellSpec
        {
            Ref = "B2", Raw = "42", Kind = ValueKind.Number, Value = "42",
            Note = "the row claims spans 5:9 and holds columns 1:2",
        });
        spec.Sheets.Add(sheet);
    }

    static void PartTargets(Corpus corpus)
    {
        var spec = New("realworld/part-targets.xlsx", "X0", "relationship targets real producers actually write");
        spec.Note("Parts are found by *relationship*, never by path convention: `xl/workbook.xml` " +
                  "is where every producer puts it and nowhere in the spec promises that. This " +
                  "file puts it somewhere else entirely, and a reader that opens the conventional " +
                  "path finds nothing.");
        spec.Note("The targets here exercise the four spellings the wild contains: a leading " +
                  "slash, a `..` segment, a backslash separator, and percent encoding. Normalise " +
                  "them all — and refuse anything that escapes the package, which is the hostile " +
                  "family's business.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook
            {
                // Not xl/. Nothing says it has to be.
                WorkbookPath = "parts/book%20one.xml",
                WorkbookTarget = "/parts/book%20one.xml",
            };
            book.SheetPaths.Clear();
            book.SheetPaths.AddRange(["parts/sheets/one.xml", "parts/sheets/two.xml", "parts/sheets/three.xml"]);
            book.SheetNames.Clear();
            book.SheetNames.AddRange(["Slash", "Dots", "Backslash"]);
            book.SheetTargets =
            [
                "/parts/sheets/one.xml",           // a leading slash: absolute within the package
                "../parts/sheets/two.xml",         // a `..` segment that resolves back to itself
                "sheets\\three.xml",               // a backslash, which is not a path separator in OPC
            ];
            book.SheetBodies.AddRange(
            [
                """<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>""",
                """<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>""",
                """<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>""",
            ]);
            book.Save(path);
        });

        spec.Note("The workbook part is at `parts/book%20one.xml` — percent-encoded, and not " +
                  "under `xl/` at all. The styles relationship beside it is relative, so it has to " +
                  "resolve against the workbook's own directory rather than against the root.");
        SheetOf(spec, "Slash", ("A1", "1", "reached through a target with a leading slash"));
        SheetOf(spec, "Dots", ("A1", "2", "reached through a target containing `..`"));
        SheetOf(spec, "Backslash", ("A1", "3", "reached through a target using a backslash"));
    }

    static void BomAndDeclaration(Corpus corpus)
    {
        var spec = New("realworld/bom-and-declaration.xlsx", "X0", "a byte order mark and an unusual XML declaration");
        spec.Note("A UTF-8 BOM on a part is common and harmless unless the reader assumes the " +
                  "first byte is `<`. So is `standalone=\"yes\"`, and so is no declaration at all. " +
                  "Every part in this file carries a BOM.");
        spec.Note("A reader that dispatches on the first byte, or that compares the first bytes " +
                  "against `<?xml`, fails on a file that opens correctly everywhere else.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook
            {
                Bom = true,
                Declaration = """<?xml version="1.0" encoding="utf-8" standalone="yes" ?>""",
            };
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1">
                    <c r="A1" t="inlineStr"><is><t>after a BOM</t></is></c>
                    <c r="B1"><v>1</v></c>
                  </row>
                </sheetData>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", null, "every part in this package begins EF BB BF"),
            ("B1", "1", null));
        spec.Sheets[0].Cells[0].Kind = ValueKind.Text;
        spec.Sheets[0].Cells[0].Value = "after a BOM";

        var second = New("realworld/no-declaration.xlsx", "X0", "parts with no XML declaration at all");
        second.Note("A declaration is optional in XML. Parts here begin with `<worksheet` and " +
                    "nothing else, which is legal and which a reader expecting a prologue will " +
                    "not survive.");

        corpus.Emit(second, path =>
        {
            var book = new RawWorkbook { Declaration = null };
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1"><v>7</v></c></row>
                </sheetData>
                """);
            book.Save(path);
        });

        SheetOf(second, "Sheet1", ("A1", "7", "the part opens with its root element"));
    }

    static void UnusualPackaging(Corpus corpus)
    {
        var spec = New("realworld/unusual-zip.xlsx", "X0", "stored entries, odd order, extra files, directory entries");
        spec.Note("Nothing requires the parts of a package to be deflated, to appear in any " +
                  "particular order, or to be the only entries in the archive. A reader that " +
                  "assumes the first entry is `[Content_Types].xml` — which a great many do, " +
                  "because Excel puts it first — fails on a perfectly ordinary file.");
        spec.Note("The extra entries here are the kind a real workflow leaves behind: a " +
                  "`.DS_Store`, an editor backup, a directory entry with a trailing slash and no " +
                  "content. None is a part, none is typed, and none is a reason to refuse.");
        spec.Note("Measured 2026-09-15: LibreOffice opens this file. Each oddity here was checked " +
                  "on its own, and none of them is a reason to refuse — which is what makes them " +
                  "the *tolerable* half. The one packaging oddity the oracle does refuse lives in " +
                  "duplicate-entry.xlsx, by itself, so that a failure on either file names its own " +
                  "cause.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
                </sheetData>
                """);
            // Stored rather than deflated: legal, and produced by anything optimising for speed.
            book.Stored.Add("xl/worksheets/sheet1.xml");
            book.Stored.Add("[Content_Types].xml");
            book.Finally = zip =>
            {
                zip.Add("xl/", "", deflate: false);
                zip.Add(".DS_Store", "not a part, and not XML either", deflate: false);
                zip.Add("xl/workbook.xml.bak", "an editor's leftovers");
                zip.Add("docProps/thumbnail.jpeg", "not really a JPEG", deflate: false);
            };
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", "1", "in a part stored rather than deflated"),
            ("B1", "2", null));

        DuplicateEntry(corpus);
    }

    static void DuplicateEntry(Corpus corpus)
    {
        var spec = New("realworld/duplicate-entry.xlsx", "X0", "two zip entries sharing one part name");
        spec.OracleOpens = false;
        spec.DuplicateEntries = true;
        spec.Note("Zip permits two entries under the same name and OPC says nothing about which " +
                  "one is the part. A reader has to choose, and the only real requirement is that " +
                  "it choose the *same* one on every run and on every machine — a file that " +
                  "imports differently twice is worse than one that fails.");
        spec.Note("Measured 2026-09-15: LibreOffice refuses this file outright, and it is the only " +
                  "packaging oddity in this family that it refuses — stored entries, directory " +
                  "entries, stray non-part files and reordered entries all open fine. So refusing " +
                  "is defensible and in good company; choosing deterministically is better.");
        spec.Note("The duplicate is `xl/styles.xml` rather than a worksheet, so that whichever " +
                  "copy wins, the cell values are the same and the fixture stays assertable. The " +
                  "second copy differs only in a font nothing references.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
                </sheetData>
                """);
            book.Finally = zip => zip.AddDuplicate("xl/styles.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="1"><font><sz val="99"/><name val="Duplicate"/></font></fonts>
                  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
                  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
                  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
                </styleSheet>
                """);
            book.Save(path);
        });

        SheetOf(spec, "Sheet1",
            ("A1", "1", "the values do not depend on which copy of the styles part wins"),
            ("B1", "2", null));
    }

    static void SheetOrder(Corpus corpus)
    {
        var spec = New("realworld/sheet-order.xlsx", "X0", "workbook order unrelated to part names or relationship order");
        spec.Note("A sheet's position comes from the order of `<sheet>` elements in the workbook " +
                  "part, and from nothing else. The part *names* here run backwards and the " +
                  "relationship ids are shuffled, so a reader that sorts by either — or that " +
                  "trusts `sheetId` — gets the order wrong in three different ways.");
        spec.Note("`sheetId` is an identifier, not an index. It need not start at 1, need not be " +
                  "contiguous, and need not ascend.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetPaths.Clear();
            book.SheetNames.Clear();
            book.SheetBodies.Clear();

            // Declared first, stored last, and given the highest sheetId.
            book.SheetPaths.AddRange(["xl/worksheets/sheet9.xml", "xl/worksheets/sheet5.xml", "xl/worksheets/sheet1.xml"]);
            book.SheetNames.AddRange(["First", "Second", "Third"]);
            book.SheetBodies.AddRange(
            [
                """<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>I am first</t></is></c></row></sheetData>""",
                """<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>I am second</t></is></c></row></sheetData>""",
                """<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>I am third</t></is></c></row></sheetData>""",
            ]);
            book.Save(path);
        });

        foreach ((string name, string text) in new[] { ("First", "I am first"), ("Second", "I am second"), ("Third", "I am third") })
        {
            var sheet = new SheetSpec { Name = name };
            sheet.Cells.Add(new CellSpec
            {
                Ref = "A1", T = "inlineStr", Kind = ValueKind.Text, Value = text,
                Note = "the sheet's position is the workbook's order, not the part name's",
            });
            spec.Sheets.Add(sheet);
        }
    }

    static void DefaultNamespace(Corpus corpus)
    {
        var spec = New("realworld/prefixed-namespace.xlsx", "X0", "the same document with prefixes instead of a default namespace");
        spec.Note("A prefix is a local choice and can be redeclared on any element, so dispatch " +
                  "has to be on (namespace URI, local name) and never on the prefix. Excel writes " +
                  "a default namespace; this file writes `<s:worksheet>` and means exactly the " +
                  "same thing.");
        spec.Note("Row 2 redeclares the *same* namespace under a second prefix mid-document and " +
                  "uses both in one row. A reader keyed on the prefix reads half a row.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();
            const string main = RawWorkbook.Transitional;
            const string rels = RawWorkbook.TransitionalRelationships;
            const string type = RawWorkbook.TransitionalRelType;

            zip.Add("[Content_Types].xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="{RawWorkbook.ContentTypesNs}">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """);

            zip.Add("_rels/.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pkg:Relationships xmlns:pkg="{RawWorkbook.PackageRelationships}">
                  <pkg:Relationship Id="rId1" Type="{type}/officeDocument" Target="xl/workbook.xml"/>
                </pkg:Relationships>
                """);

            zip.Add("xl/workbook.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <s:workbook xmlns:s="{main}" xmlns:rel="{rels}">
                  <s:sheets>
                    <s:sheet name="Prefixed" sheetId="1" rel:id="rId1"/>
                  </s:sheets>
                </s:workbook>
                """);

            zip.Add("xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{RawWorkbook.PackageRelationships}">
                  <Relationship Id="rId1" Type="{type}/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """);

            zip.Add("xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <s:worksheet xmlns:s="{main}">
                  <s:sheetData>
                    <s:row r="1">
                      <s:c r="A1"><s:v>1</s:v></s:c>
                      <s:c r="B1" t="inlineStr"><s:is><s:t>prefixed</s:t></s:is></s:c>
                    </s:row>
                    <s:row r="2" xmlns:other="{main}">
                      <s:c r="A2"><s:v>2</s:v></s:c>
                      <other:c r="B2"><other:v>3</other:v></other:c>
                    </s:row>
                  </s:sheetData>
                </s:worksheet>
                """);

            zip.Save(path);
        });

        var sheet = new SheetSpec { Name = "Prefixed" };
        sheet.Cells.Add(new CellSpec { Ref = "A1", Raw = "1", Kind = ValueKind.Number, Value = "1", Note = "under the prefix `s`" });
        sheet.Cells.Add(new CellSpec { Ref = "B1", T = "inlineStr", Kind = ValueKind.Text, Value = "prefixed" });
        sheet.Cells.Add(new CellSpec { Ref = "A2", Raw = "2", Kind = ValueKind.Number, Value = "2" });
        sheet.Cells.Add(new CellSpec
        {
            Ref = "B2", Raw = "3", Kind = ValueKind.Number, Value = "3",
            Note = "the prefix `other` bound to the same URI, declared on the row — the same element " +
                   "type as its neighbour, spelled differently",
        });
        spec.Sheets.Add(sheet);
    }

    // --- shared plumbing ---

    /// <summary>
    /// Record a sheet's expectations for a fixture whose bytes were written as text. Numeric
    /// by default; a caller adjusts the entries that are not.
    /// </summary>
    static void SheetOf(FixtureSpec spec, string name, params (string Ref, string? Value, string? Note)[] cells)
    {
        var sheet = new SheetSpec { Name = name };
        foreach ((string reference, string? value, string? note) in cells)
        {
            sheet.Cells.Add(new CellSpec
            {
                Ref = reference,
                Raw = value,
                Kind = value is null ? ValueKind.Text : ValueKind.Number,
                Value = value,
                Note = note,
            });
        }

        spec.Sheets.Add(sheet);
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "realworld", Milestone = milestone, Covers = covers };
}
