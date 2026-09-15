using System.Globalization;
using System.Text;
using OoxmlGen.Raw;

namespace OoxmlGen.Families;

/// <summary>
/// X1 — one large workbook, for the timing test.
///
/// Excel files with a million rows exist. The claim worth testing is that a reader streams
/// and bounds what it materialises rather than building the whole sheet in memory first, and
/// that claim needs a file rather than an assumption.
/// </summary>
public static class Scale
{
    const int Rows = 100_000;
    const int Columns = 5;

    public static void Generate(Corpus corpus)
    {
        Large(corpus);
        WideAndSparse(corpus);
    }

    static void Large(Corpus corpus)
    {
        var spec = New("scale/large-sheet.xlsx", "X1", $"{Rows * Columns:N0} cells in one sheet");
        spec.Note($"{Rows:N0} rows of {Columns} columns: numbers, a shared string, a date and a " +
                  "formula with a cached value, so the cost being measured is the whole read path " +
                  "and not just integer parsing.");
        spec.Note("This is the timing fixture. It is generated rather than vendored because a " +
                  "few megabytes of repetitive XML has no business in a git history, and because " +
                  "the row count is a constant one edit away from being a million.");
        spec.Note("Only the first and last rows are asserted. The point of the file is how long " +
                  "it takes and how much memory it costs, not what is in the middle of it — and " +
                  "a manifest listing half a million cells would be larger than the workbook.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook
            {
                SharedStrings = $"""
                    <sst xmlns="{RawWorkbook.Transitional}" count="{Rows}" uniqueCount="4">
                      <si><t>north</t></si>
                      <si><t>south</t></si>
                      <si><t>east</t></si>
                      <si><t>west</t></si>
                    </sst>
                    """,
            };
            book.SheetBodies.Add("<sheetData/>");
            book.Finally = zip => zip.AddStreamed("xl/worksheets/sheet1.xml", stream =>
            {
                var writer = new StreamWriter(stream, new UTF8Encoding(false), 1 << 16);
                writer.Write($"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
                writer.Write($"""<worksheet xmlns="{RawWorkbook.Transitional}">""");
                writer.Write($"""<dimension ref="A1:E{Rows}"/>""");
                writer.Write("<sheetData>");

                for (int row = 1; row <= Rows; row++)
                {
                    double value = row * 1.5;
                    writer.Write($"""<row r="{row}">""");
                    writer.Write($"""<c r="A{row}"><v>{row}</v></c>""");
                    writer.Write($"""<c r="B{row}"><v>{value.ToString("R", CultureInfo.InvariantCulture)}</v></c>""");
                    writer.Write($"""<c r="C{row}" t="s"><v>{row % 4}</v></c>""");
                    writer.Write($"""<c r="D{row}" s="1"><v>{45000 + row % 365}</v></c>""");
                    writer.Write($"""<c r="E{row}"><f>A{row}*B{row}</f><v>{(row * value).ToString("R", CultureInfo.InvariantCulture)}</v></c>""");
                    writer.Write("</row>");
                }

                writer.Write("</sheetData></worksheet>");
                writer.Flush();
            });
            book.Save(path);
        });

        var sheet = new SheetSpec { Name = "Sheet1" };
        sheet.Cells.Add(new CellSpec
        {
            Ref = "A1", Raw = "1", Kind = ValueKind.Number, Value = "1",
            Note = $"the first of {Rows * Columns:N0} cells; the rows between here and the last are " +
                   "not asserted",
        });
        sheet.Cells.Add(new CellSpec { Ref = "B1", Raw = "1.5", Kind = ValueKind.Number, Value = "1.5" });
        sheet.Cells.Add(new CellSpec { Ref = "C1", T = "s", Kind = ValueKind.Text, Value = "south" });
        sheet.Cells.Add(new CellSpec
        {
            Ref = "E1", Raw = "1.5", Kind = ValueKind.Number, Value = "1.5",
            Excel = "A1*B1", Formula = "[.A1]*[.B1]",
        });
        sheet.Cells.Add(new CellSpec
        {
            Ref = $"A{Rows}", Raw = Rows.ToString(), Kind = ValueKind.Number, Value = Rows.ToString(),
            Note = "the last row: a reader that stops early fails here rather than anywhere visible",
        });
        sheet.Cells.Add(new CellSpec
        {
            Ref = $"B{Rows}", Raw = (Rows * 1.5).ToString("R", CultureInfo.InvariantCulture),
            Kind = ValueKind.Number, Value = (Rows * 1.5).ToString("R", CultureInfo.InvariantCulture),
        });
        spec.Sheets.Add(sheet);
    }

    static void WideAndSparse(Corpus corpus)
    {
        var spec = New("scale/wide-and-sparse.xlsx", "X1", "cells at the far corners of a sheet, and nothing between");
        spec.Note("Six cells, at A1 and at the last row and column Excel has. A reader that " +
                  "materialises the rectangle between them builds 17 billion cells for a file of " +
                  "a few hundred bytes — the same failure as trusting `dimension`, reached from " +
                  "the other direction, because here the dimension is honest.");
        spec.Note("XFD is column 16384 and 1048576 is the last row. Both are the real limits, so " +
                  "a reader that rejects them as out of range rejects a legal file.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("""
                <dimension ref="A1:XFD1048576"/>
                <sheetData>
                  <row r="1">
                    <c r="A1"><v>1</v></c>
                    <c r="XFD1"><v>2</v></c>
                  </row>
                  <row r="500000">
                    <c r="A500000"><v>3</v></c>
                    <c r="XFD500000"><v>4</v></c>
                  </row>
                  <row r="1048576">
                    <c r="A1048576"><v>5</v></c>
                    <c r="XFD1048576"><v>6</v></c>
                  </row>
                </sheetData>
                """);
            book.Save(path);
        });

        var sheet = new SheetSpec { Name = "Sheet1" };
        var corners = new (string Ref, string Value, string? Note)[]
        {
            ("A1", "1", "the only cell a naive reader is sure to find"),
            ("XFD1", "2", "column 16384, the last one Excel has"),
            ("A500000", "3", null),
            ("XFD500000", "4", null),
            ("A1048576", "5", "row 1048576, the last one Excel has"),
            ("XFD1048576", "6", "the far corner: legal, and the whole sheet away from A1"),
        };

        foreach ((string reference, string value, string? note) in corners)
        {
            sheet.Cells.Add(new CellSpec
            {
                Ref = reference, Raw = value, Kind = ValueKind.Number, Value = value, Note = note,
            });
        }

        spec.Sheets.Add(sheet);
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "scale", Milestone = milestone, Covers = covers };
}
