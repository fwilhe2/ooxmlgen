using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X2 — the expression translator. Excel's A1 syntax in, OpenFormula out, and the classes
/// that are excluded rather than translated.
///
/// The expected spellings here follow grind's own table in <c>doc/xlsx-import.md</c> Part II
/// §3. Where that table is silent the expectation is left null rather than guessed: a
/// fixture that asserts an invented canonical form tests the fixture, not the importer.
/// </summary>
public static class Formulas
{
    public static void Generate(Corpus corpus)
    {
        References(corpus);
        Operators(corpus);
        Functions(corpus);
        SharedGroups(corpus);
        ExtendedFunctions(corpus);
        NoCachedValue(corpus);
        ExcludedClasses(corpus);
        Errors(corpus);
        DivergentSemantics(corpus);
    }

    /// <summary>Every shape a reference takes, including the ones that name another sheet.</summary>
    static void References(Corpus corpus)
    {
        var spec = New("formulas/references.xlsx", "X2", "references: absolute, mixed, whole-track, cross-sheet, 3-D");
        spec.Note("Sheet names are chosen to be awkward on purpose: one has a space, one has an " +
                  "apostrophe that has to be doubled inside the quotes on both sides of the " +
                  "translation.");
        spec.Note("`formula` follows grind's own table, which spells a cross-sheet reference " +
                  "`[Data.A1]`. `oracle` is what LibreOffice writes for the same cell, measured on " +
                  "2026-09-14: `[$Data.A1]`, with a `$` marking the sheet name absolute. Excel has " +
                  "no sheet-relative form to distinguish, so both readings are defensible — but a " +
                  "comparison against the oracle has to know which one it is looking at.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab data = book.Sheet("Data");
            Tab spaced = book.Sheet("My Sheet");
            Tab quoted = book.Sheet("O'Brien");
            Tab third = book.Sheet("Sheet3");

            foreach (Tab sheet in new[] { data, spaced, quoted, third })
                for (int row = 1; row <= 3; row++)
                    sheet.Number($"A{row}", row * 10).Number($"B{row}", row * 100);

            Tab tab = book.Sheet("Refs");
            var cases = new (string Label, string Excel, string? Expect, string Oracle, string Cached, string? Note)[]
            {
                ("relative", "Data!A1", "[Data.A1]", "[$Data.A1]", "10", null),
                ("absolute", "Data!$A$1", "[Data.$A$1]", "[$Data.$A$1]", "10", null),
                ("column-absolute", "Data!$A1", "[Data.$A1]", "[$Data.$A1]", "10", null),
                ("row-absolute", "Data!A$1", "[Data.A$1]", "[$Data.A$1]", "10", null),
                ("range", "SUM(Data!A1:B3)", "SUM([Data.A1:.B3])", "SUM([$Data.A1:.B3])", "660",
                 "the end of a range drops the repeated sheet name — in both spellings"),
                ("whole column", "SUM(Data!A:A)", "SUM([Data.A:.A])", "SUM([$Data.A:.A])", "60", null),
                ("whole row", "SUM(Data!1:1)", "SUM([Data.1:.1])", "SUM([$Data.1:.1])", "110", null),
                ("sheet with a space", "'My Sheet'!A2", "['My Sheet'.A2]", "[$'My Sheet'.A2]", "20", null),
                ("sheet with an apostrophe", "'O''Brien'!A3", "['O''Brien'.A3]", "[$'O''Brien'.A3]", "30",
                 "the apostrophe is doubled inside the quotes — in both spellings"),
                ("3-D sum", "SUM(Data:Sheet3!A1)", "SUM([Data.A1:Sheet3.A1])", "SUM([$Data.A1:$Sheet3.A1])", "40",
                 "a cuboid: the same cell across a run of sheets"),
                ("3-D range", "SUM(Data:Sheet3!A1:A2)", null, "SUM([$Data.A1:Sheet3.A2])", "120",
                 "grind's table does not state this one, so `formula` asserts nothing. The oracle's " +
                 "answer is odd in its own right: the second sheet name loses the `$` the first kept"),
            };

            int r = 1;
            foreach ((string label, string excel, string? expect, string oracle, string cached, string? note) in cases)
            {
                tab.Text($"A{r}", label).Formula($"B{r}", excel, cached, expect).Oracle(oracle);
                if (note is not null) tab.Expect(note: note);
                r++;
            }
        }));
    }

    /// <summary>Operators, and the two places precedence surprises people.</summary>
    static void Operators(Corpus corpus)
    {
        var spec = New("formulas/operators.xlsx", "X2", "arithmetic, comparison, concatenation, precedence");
        spec.Note("`-2^2` is 4 in Excel, because prefix minus binds tighter than `^`. OpenFormula " +
                  "§5.5 agrees, so this is one of the few places the surprise carries across " +
                  "unchanged — and one worth a fixture precisely because it looks like a bug.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Operators");
            tab.Number("A1", 6).Number("A2", 4).Text("A3", "left").Text("A4", "right");

            var cases = new (string Excel, string? Expect, string Cached, CellValues? Type, string? Note)[]
            {
                ("A1+A2", "[.A1]+[.A2]", "10", null, null),
                ("A1-A2", "[.A1]-[.A2]", "2", null, null),
                ("A1*A2", "[.A1]*[.A2]", "24", null, null),
                ("A1/A2", "[.A1]/[.A2]", "1.5", null, null),
                ("A1^A2", "[.A1]^[.A2]", "1296", null, null),
                ("-A1", "-[.A1]", "-6", null, null),
                ("+A1", "+[.A1]", "6", null, "unary plus: legal, and almost never written"),
                ("A1%", "[.A1]%", "0.06", null, "postfix percent"),
                ("-2^2", "-2^2", "4", null, "prefix minus above `^` — 4, not -4"),
                ("(-2)^2", "(-2)^2", "4", null, "the parenthesised form, for contrast"),
                ("2^3^2", "2^3^2", "64", null, "`^` is left-associative in Excel: (2^3)^2, not 2^(3^2)"),
                ("A3&A4", "[.A3]&[.A4]", "leftright", CellValues.String, null),
                ("A1=A2", "[.A1]=[.A2]", "0", CellValues.Boolean, null),
                ("A1<>A2", "[.A1]<>[.A2]", "1", CellValues.Boolean, null),
                ("A1>A2", "[.A1]>[.A2]", "1", CellValues.Boolean, null),
                ("A1<=A2", "[.A1]<=[.A2]", "0", CellValues.Boolean, null),
            };

            int r = 1;
            foreach ((string excel, string? expect, string cached, CellValues? type, string? note) in cases)
            {
                ValueKind kind = type switch
                {
                    null => ValueKind.Number,
                    var t when t.Value == CellValues.Boolean => ValueKind.Bool,
                    _ => ValueKind.Text,
                };

                string? value = kind == ValueKind.Bool ? cached == "1" ? "true" : "false" : cached;
                tab.Formula($"C{r}", excel, cached, expect, type, kind: kind).Expect(value: value);
                if (note is not null) tab.Expect(note: note);
                r++;
            }
        }));
    }

    /// <summary>Functions this build implements, so a recalculation has something to agree with.</summary>
    static void Functions(Corpus corpus)
    {
        var spec = New("formulas/functions.xlsx", "X2", "functions grind implements, with Excel's cached values");
        spec.Note("Nothing here should be reported unknown. The point of the file is the other " +
                  "loop: recalculating it must reproduce the values Excel cached, and where it " +
                  "does not, the two implementations disagree about the function rather than " +
                  "about the import.");
        spec.Note("`TRUE` and `FALSE` are bare words in Excel and functions in OpenFormula — " +
                  "§6.15 — so they gain their parentheses on the way across.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Functions");
            for (int row = 1; row <= 5; row++) tab.Number($"A{row}", row * 2);
            tab.Text("B1", "hello").Text("B2", "WORLD").Text("B3", "  pad  ");

            var cases = new (string Excel, string? Expect, string Cached, CellValues? Type)[]
            {
                ("SUM(A1:A5)", "SUM([.A1:.A5])", "30", null),
                ("AVERAGE(A1:A5)", "AVERAGE([.A1:.A5])", "6", null),
                ("MIN(A1:A5)", "MIN([.A1:.A5])", "2", null),
                ("MAX(A1:A5)", "MAX([.A1:.A5])", "10", null),
                ("COUNT(A1:A5)", "COUNT([.A1:.A5])", "5", null),
                ("COUNTA(A1:B5)", "COUNTA([.A1:.B5])", "8", null),
                ("ROUND(A1/3,2)", "ROUND([.A1]/3;2)", "0.67", null),
                ("ABS(-A1)", "ABS(-[.A1])", "2", null),
                ("INT(7.9)", "INT(7.9)", "7", null),
                ("MOD(7,3)", "MOD(7;3)", "1", null),
                ("POWER(2,10)", "POWER(2;10)", "1024", null),
                ("SQRT(144)", "SQRT(144)", "12", null),
                ("IF(A1>1,\"yes\",\"no\")", "IF([.A1]>1;\"yes\";\"no\")", "yes", CellValues.String),
                ("AND(A1>1,A2>1)", "AND([.A1]>1;[.A2]>1)", "1", CellValues.Boolean),
                ("OR(A1>100,A2>1)", "OR([.A1]>100;[.A2]>1)", "1", CellValues.Boolean),
                ("NOT(A1>100)", "NOT([.A1]>100)", "1", CellValues.Boolean),
                ("TRUE", "TRUE()", "1", CellValues.Boolean),
                ("FALSE", "FALSE()", "0", CellValues.Boolean),
                ("LEN(B1)", "LEN([.B1])", "5", null),
                ("UPPER(B1)", "UPPER([.B1])", "HELLO", CellValues.String),
                ("LOWER(B2)", "LOWER([.B2])", "world", CellValues.String),
                ("TRIM(B3)", "TRIM([.B3])", "pad", CellValues.String),
                ("LEFT(B1,2)", "LEFT([.B1];2)", "he", CellValues.String),
                ("RIGHT(B1,2)", "RIGHT([.B1];2)", "lo", CellValues.String),
                ("MID(B1,2,3)", "MID([.B1];2;3)", "ell", CellValues.String),
                ("VLOOKUP(4,A1:A5,1,FALSE)", "VLOOKUP(4;[.A1:.A5];1;FALSE())", "4", null),
                ("DATE(2024,3,17)", "DATE(2024;3;17)", "45368", null),
                ("SUMIF(A1:A5,\">4\")", "SUMIF([.A1:.A5];\">4\")", "24", null),
            };

            int r = 1;
            foreach ((string excel, string? expect, string cached, CellValues? type) in cases)
            {
                ValueKind kind = type switch
                {
                    null => ValueKind.Number,
                    var t when t.Value == CellValues.Boolean => ValueKind.Bool,
                    _ => ValueKind.Text,
                };

                string value = kind == ValueKind.Bool ? cached == "1" ? "true" : "false" : cached;
                tab.Formula($"D{r}", excel, cached, expect, type, kind: kind).Expect(value: value);
                r++;
            }
        }));
    }

    /// <summary>Shared-formula groups, which is how Excel stores a filled-down column.</summary>
    static void SharedGroups(Corpus corpus)
    {
        var spec = New("formulas/shared-groups.xlsx", "X2", "t=\"shared\" groups, and one that shifts off the sheet");
        spec.Note("A follower carries `<f t=\"shared\" si=\"n\"/>` and nothing else. Its expression " +
                  "is the master's with every relative axis shifted by the offset between the two " +
                  "cells — which is what grind's existing formula::shift already does for fill.");
        spec.Note("Group 2 is the case that catches a naive implementation: its master sits at B2 " +
                  "and refers to A1, so the follower at B1 would refer to A0. A reference that " +
                  "leaves the sheet becomes #REF!, exactly as it does when a row is deleted.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Shared");

            for (int row = 1; row <= 6; row++) tab.Number($"A{row}", row);

            // Group 0: the ordinary filled-down column.
            tab.SharedFormulaMaster("B1", 0, "B1:B6", "A1*2", "2", "[.A1]*2");
            for (int row = 2; row <= 6; row++)
                tab.SharedFormulaMember($"B{row}", 0, (row * 2).ToString(), $"[.A{row}]*2");

            // Group 1: a two-dimensional group, so both axes shift.
            tab.SharedFormulaMaster("C2", 1, "C2:D3", "A2+B2", "4", "[.A2]+[.B2]");
            tab.SharedFormulaMember("D2", 1, "6", "[.B2]+[.C2]");
            tab.SharedFormulaMember("C3", 1, "9", "[.A3]+[.B3]");
            tab.SharedFormulaMember("D3", 1, "12", "[.B3]+[.C3]");

            // Group 2: shifting the follower up takes the reference off the top of the sheet.
            tab.SharedFormulaMaster("F2", 2, "F1:F2", "A1*10", "10", "[.A1]*10");
            tab.SharedFormulaMember("F1", 2, null, "[#REF!]*10")
               .Expect(note: "shifted up one row from the master: A1 becomes A0, which does not exist. " +
                             "Only the *reference* breaks — the rest of the expression survives, which " +
                             "is what `[#REF!]*10` says and what a whole-formula #REF! would not. " +
                             "Measured against LibreOffice on 2026-09-14.");

            // Group 3: absolute references do not shift at all.
            tab.SharedFormulaMaster("G1", 3, "G1:G3", "$A$1+A1", "2", "[.$A$1]+[.A1]");
            tab.SharedFormulaMember("G2", 3, "3", "[.$A$1]+[.A2]")
               .Expect(note: "the absolute half stays put; only the relative half moves");
            tab.SharedFormulaMember("G3", 3, "4", "[.$A$1]+[.A3]");
        }));
    }

    /// <summary>The `_xlfn.` prefix, and the functions that carry it.</summary>
    static void ExtendedFunctions(Corpus corpus)
    {
        var spec = New("formulas/xlfn.xlsx", "X2", "_xlfn.-prefixed functions and the implicit-intersection marker");
        spec.Unknown("XLOOKUP", "TEXTJOIN", "IFS", "SWITCH", "CONCAT", "MAXIFS");
        spec.Note("Excel writes functions added after 2007 with an `_xlfn.` prefix so older " +
                  "versions round-trip them without pretending to understand them. The prefix is " +
                  "stripped on import; the bare name is then reported unknown, which is a different " +
                  "and honest statement.");
        spec.Note("`_xlfn.SINGLE(A1)` is the implicit-intersection marker that modern Excel shows " +
                  "as `@A1`. It is dropped, leaving the reference behind.");
        spec.Note("The oracle takes the other road entirely: LibreOffice maps these onto ODF's " +
                  "namespaced-function convention, `COM.MICROSOFT.XLOOKUP`, rather than stripping " +
                  "the prefix and admitting it does not know the function. Measured 2026-09-14. " +
                  "Neither is wrong, and a comparison that does not expect the difference reads " +
                  "every one of these cells as a failure.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Extended");
            for (int row = 1; row <= 4; row++) tab.Number($"A{row}", row).Text($"B{row}", $"row {row}");

            var cases = new (string Excel, string? Expect, string Oracle, string Cached, CellValues? Type, string? Note)[]
            {
                ("_xlfn.XLOOKUP(3,A1:A4,B1:B4)", "XLOOKUP(3;[.A1:.A4];[.B1:.B4])",
                 "COM.MICROSOFT.XLOOKUP(3;[.A1:.A4];[.B1:.B4])", "row 3", CellValues.String,
                 "the prefix goes; the name is then unknown to this build"),
                ("_xlfn.TEXTJOIN(\",\",TRUE,B1:B4)", "TEXTJOIN(\",\";TRUE();[.B1:.B4])",
                 "COM.MICROSOFT.TEXTJOIN(\",\";TRUE();[.B1:.B4])",
                 "row 1,row 2,row 3,row 4", CellValues.String, null),
                ("_xlfn.IFS(A1>3,\"big\",A1>0,\"small\")", "IFS([.A1]>3;\"big\";[.A1]>0;\"small\")",
                 "COM.MICROSOFT.IFS([.A1]>3;\"big\";[.A1]>0;\"small\")",
                 "small", CellValues.String, null),
                ("_xlfn.SWITCH(A2,1,\"one\",2,\"two\",\"other\")",
                 "SWITCH([.A2];1;\"one\";2;\"two\";\"other\")",
                 "COM.MICROSOFT.SWITCH([.A2];1;\"one\";2;\"two\";\"other\")", "two", CellValues.String, null),
                ("_xlfn.CONCAT(B1,B2)", "CONCAT([.B1];[.B2])", "COM.MICROSOFT.CONCAT([.B1];[.B2])",
                 "row 1row 2", CellValues.String, null),
                ("_xlfn.MAXIFS(A1:A4,A1:A4,\"<4\")", "MAXIFS([.A1:.A4];[.A1:.A4];\"<4\")",
                 "COM.MICROSOFT.MAXIFS([.A1:.A4];[.A1:.A4];\"<4\")", "3", null, null),
                ("SUM(_xlfn.SINGLE(A1:A4))", "SUM([.A1:.A4])", "SUM(_xlfn.single([.A1:.A4]))", "1", null,
                 "implicit intersection: the marker is dropped and the reference kept. The oracle " +
                 "keeps the marker instead, lowercased, which is the one case here where it carries " +
                 "a construct forward that has no meaning on the other side"),
            };

            int r = 1;
            foreach ((string excel, string? expect, string oracle, string cached, CellValues? type, string? note) in cases)
            {
                ValueKind kind = type is null ? ValueKind.Number : ValueKind.Text;
                tab.Formula($"D{r}", excel, cached, expect, type, kind: kind).Oracle(oracle);
                if (note is not null) tab.Expect(note: note);
                r++;
            }
        }));
    }

    /// <summary>A formula with no cached value, which is what a producer that does not evaluate writes.</summary>
    static void NoCachedValue(Corpus corpus)
    {
        var spec = New("formulas/no-cached-value.xlsx", "X2", "<f> with no <v> — the only formula cell with no number");
        spec.Note("Apache POI and SheetJS write this by default: they do not evaluate, so there is " +
                  "nothing to cache. Given that nothing is evaluated on import either, these are the " +
                  "only cells that arrive with a formula and no value — carry the formula, leave the " +
                  "value empty, count them, and let `sheet recalc` be the thing that fills them in.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Uncached");
            tab.Number("A1", 3).Number("A2", 4);

            tab.Formula("B1", "A1+A2", null, "[.A1]+[.A2]")
               .ExpectNoValue("no <v> at all: the formula is carried, the value is empty");
            tab.Formula("B2", "SUM(A1:A2)", null, "SUM([.A1:.A2])")
               .ExpectNoValue("likewise");
            tab.Formula("B3", "A1&\"x\"", null, "[.A1]&\"x\"")
               .ExpectNoValue("a text-valued formula with nothing cached: not even `t` says what it is");
            tab.Number("C1", 7).Expect(note: "a plain value in the same sheet, to prove the file is otherwise ordinary");
        }));
    }

    /// <summary>The constructs that are excluded by design rather than unimplemented.</summary>
    static void ExcludedClasses(Corpus corpus)
    {
        var spec = New("formulas/excluded-classes.xlsx", "X2", "array formulas, inline arrays, intersection, union");
        spec.Drops(Dropped.ArrayFormula, 3);
        spec.Note("Each of these is a *class*, named in grind's exclusion list, not a function " +
                  "that happens to be missing. The cached value survives in every case; it is the " +
                  "formula that does not.");
        spec.Note("Excel's intersection operator is a space and its union operator is a comma — " +
                  "the same comma that separates arguments, disambiguated by position. Both are " +
                  "outside the Small Group.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Excluded");
            for (int row = 1; row <= 3; row++)
                for (int col = 1; col <= 3; col++)
                    tab.Number(A1.Name(col, row), row * 10 + col);

            tab.ArrayFormula("E1", "E1:E1", "SUM(A1:C3*2)", "198")
               .Expect(note: "a single-cell array formula: legal, and still an array formula")
               .Oracle("SUM([.A1:.C3]*2)");
            tab.ArrayFormula("E2", "E2:E4", "A1:A3*2", "22")
               .Expect(note: "the master of a multi-cell array; E3 and E4 belong to the same range")
               .Oracle("[.A1:.A3]*2");
            tab.Number("E3", 42).Expect(note: "inside the array's ref but written as a plain cached value");
            tab.Number("E4", 62);

            tab.Formula("F1", "SUM({1,2;3,4})", "10", null)
               .Expect(note: "an inline array constant — §5.13, excluded by §2.3.2. ODF separates " +
                             "array columns with `;` and rows with `|`, so Excel's `{1,2;3,4}` is " +
                             "`{1;2|3;4}` there: both separators change meaning at once")
               .Expect(kind: ValueKind.Number)
               .Oracle("SUM({1;2|3;4})");
            tab.ArrayFormula("F2", "F2:F2", "SUM({1,2;3,4})", "10")
               .Oracle("SUM({1;2|3;4})");

            tab.Formula("G1", "SUM(A1:C2 B1:B3)", "22", null)
               .Expect(note: "intersection: the space operator. OpenFormula spells it `!` and the " +
                             "oracle does emit it, so this is expressible in ODF and still outside " +
                             "the Small Group — dropped by choice rather than by inability")
               .Oracle("SUM([.A1:.C2]![.B1:.B3])");
            tab.Formula("G2", "SUM((A1:A2,C1:C2))", "70", null)
               .Expect(note: "union: the comma operator, inside an extra pair of parentheses so it " +
                             "is not read as two arguments. ODF spells it `~`")
               .Oracle("SUM(([.A1:.A2]~[.C1:.C2]))");
        }));
    }

    /// <summary>Formulas whose cached value is an error.</summary>
    static void Errors(Corpus corpus)
    {
        var spec = New("formulas/errors.xlsx", "X2", "formulas that cached an error rather than a number");
        spec.Note("An error is a value, not a failure to have one: it has a `t=\"e\"` and a name, " +
                  "and the name is the same in both formula languages. A reader that treats `t=\"e\"` " +
                  "as a parse problem loses a cell that Excel considers perfectly well-formed.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Errors");
            tab.Number("A1", 1).Number("A2", 0).Text("A3", "text");

            var cases = new (string Excel, string? Expect, string? Oracle, string Error, string? Note)[]
            {
                ("A1/A2", "[.A1]/[.A2]", null, "#DIV/0!", null),
                ("NA()", "NA()", null, "#N/A", null),
                ("A1+A3", "[.A1]+[.A3]", null, "#VALUE!", "text in an arithmetic operand"),
                ("NOSUCHFUNCTION(A1)", "NOSUCHFUNCTION([.A1])", "nosuchfunction([.A1])", "#NAME?",
                 "an unknown name is carried, not rejected. The oracle lowercases it, which is a " +
                 "difference in spelling rather than in meaning"),
                ("SQRT(-1)", "SQRT(-1)", null, "#NUM!", null),
                ("SUM(A1:A2 B5:B6)", null, "SUM([.A1:.A2]![.B5:.B6])", "#NULL!",
                 "two ranges that do not overlap: the intersection is empty, which is what #NULL! " +
                 "means and almost nobody writes on purpose"),
            };

            int r = 1;
            foreach ((string excel, string? expect, string? oracle, string error, string? note) in cases)
            {
                tab.Formula($"C{r}", excel, error, expect, CellValues.Error, kind: ValueKind.Error)
                   .Expect(value: error);
                if (oracle is not null) tab.Oracle(oracle);
                if (note is not null) tab.Expect(note: note);
                r++;
            }

            spec.Unknown("NOSUCHFUNCTION");
        }));
    }

    /// <summary>Functions whose name carries across but whose meaning does not.</summary>
    static void DivergentSemantics(Corpus corpus)
    {
        var spec = New("formulas/semantics-differ.xlsx", "X2",
                       "same name, different rule: the cases that only show up on recalculation");
        spec.Unknown("CEILING", "FLOOR", "ROUNDUP", "ROUNDDOWN");
        spec.Note("This fixture asserts nothing about who is right. Its job is to be the substrate " +
                  "for a measurement: import it, recalculate it, and every cell whose value moves " +
                  "names a function whose ODF rule differs from Excel's. That list belongs in " +
                  "doc/xlsx-format.md, and it cannot be written without a file like this one.");
        spec.Note("CEILING and FLOOR are additionally not implemented by this build at all, so they " +
                  "arrive as unknown functions with Excel's value intact — the two failure modes " +
                  "stacked, and worth separating in a report.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Divergent");

            var cases = new (string Excel, string? Expect, string Cached, string? Note)[]
            {
                ("MOD(-3,2)", "MOD(-3;2)", "1", "negative dividend: floored, not truncated"),
                ("MOD(3,-2)", "MOD(3;-2)", "-1", "negative divisor: the sign follows the divisor"),
                ("MOD(-3,-2)", "MOD(-3;-2)", "-1", "both negative"),
                ("ROUND(0.5,0)", "ROUND(0.5;0)", "1", "Excel rounds halves away from zero, not to even"),
                ("ROUND(1.5,0)", "ROUND(1.5;0)", "2", "banker's rounding would give 2 here too — no signal"),
                ("ROUND(2.5,0)", "ROUND(2.5;0)", "3", "banker's rounding would give 2: this is the cell that tells them apart"),
                ("ROUND(-2.5,0)", "ROUND(-2.5;0)", "-3", "away from zero in the negative direction as well"),
                ("ROUND(1.005,2)", "ROUND(1.005;2)", "1.01", "1.005 is not exactly representable; the answer depends on how the rounding is staged"),
                ("CEILING(-4.5,1)", "CEILING(-4.5;1)", "-4", "Excel's CEILING with a negative operand"),
                ("CEILING(4.5,-1)", "CEILING(4.5;-1)", "#NUM!", "opposite signs: an error rather than an answer"),
                ("FLOOR(-4.5,1)", "FLOOR(-4.5;1)", "-5", "FLOOR's mirror of the case above"),
                ("INT(-2.5)", "INT(-2.5)", "-3", "INT floors; TRUNC truncates. The two differ only for negatives"),
                ("TRUNC(-2.5)", "TRUNC(-2.5)", "-2", "and this is the cell that shows it"),
            };

            int r = 1;
            foreach ((string excel, string? expect, string cached, string? note) in cases)
            {
                bool isError = cached.StartsWith('#');
                tab.Text($"A{r}", excel);
                tab.Formula($"B{r}", excel, cached, expect,
                            isError ? CellValues.Error : null,
                            kind: isError ? ValueKind.Error : ValueKind.Number);
                if (note is not null) tab.Expect(note: note);
                r++;
            }
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "formulas", Milestone = milestone, Covers = covers };
}
