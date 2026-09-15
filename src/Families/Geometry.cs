using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X4 — column widths, row heights, and the tracks that are hidden or grouped. The one part
/// of the appearance that is a measurement rather than a mapping: Excel states a column width
/// in characters of the Normal font, and turning that into a length needs a constant nobody
/// can look up.
/// </summary>
public static class Geometry
{
    public static void Generate(Corpus corpus)
    {
        ColumnWidths(corpus);
        RowHeights(corpus);
        Outlines(corpus);
        Panes(corpus);
    }

    static void ColumnWidths(Corpus corpus)
    {
        var spec = New("geometry/columns.xlsx", "X4", "col widths, spans, hidden columns, and the default");
        spec.Note("`width` is counted in characters of the Normal font's widest digit, per " +
                  "ECMA-376 §18.3.1.13: width = trunc((chars * MDW + 5) / MDW * 256) / 256. For " +
                  "Calibri 11 the maximum digit width is 7px, which is where 8.43 characters ≈ 64px " +
                  "comes from. The constant is a property of the font, so it is a measurement " +
                  "against the oracle rather than a number to copy out of the prose.");
        spec.Note("A `<col>` covers the range min..max, not one column. Producers differ wildly " +
                  "here: Excel writes one element per run of equal widths, LibreOffice writes one " +
                  "per column. Both are legal and a reader has to expand either.");
        spec.Note("`customWidth=\"0\"` with a `width` present means the width is Excel's own " +
                  "calculation rather than the user's choice — worth carrying, and not the same " +
                  "claim as a width the author set.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Columns");

            tab.FormatProperties = new SheetFormatProperties
            {
                DefaultColumnWidth = 8.43,
                DefaultRowHeight = 15,
            };

            tab.Columns = new Columns(
                new Column { Min = 1, Max = 1, Width = 4, CustomWidth = true },
                new Column { Min = 2, Max = 2, Width = 8.43, CustomWidth = true },
                new Column { Min = 3, Max = 3, Width = 20, CustomWidth = true },
                new Column { Min = 4, Max = 4, Width = 50.5, CustomWidth = true },
                // One element covering a run of five, which is how Excel writes repeated widths.
                new Column { Min = 5, Max = 9, Width = 12.75, CustomWidth = true },
                new Column { Min = 10, Max = 10, Width = 0, Hidden = true, CustomWidth = true },
                new Column { Min = 11, Max = 11, Width = 15, Hidden = true, CustomWidth = true },
                new Column { Min = 12, Max = 12, Width = 11.5, BestFit = true, CustomWidth = true },
                new Column { Min = 13, Max = 13, Width = 9.140625 },
                // A run to the last column in the sheet: a claim about 16371 columns, and not
                // a reason to materialise any of them.
                new Column { Min = 14, Max = 16384, Width = 8.43 });

            tab.Text("A1", "4").Text("B1", "8.43").Text("C1", "20").Text("D1", "50.5");
            tab.Text("E1", "12.75").Text("F1", "run").Text("G1", "run").Text("H1", "run").Text("I1", "run");
            tab.Text("J1", "hidden, width 0").Text("K1", "hidden, width 15");
            tab.Text("L1", "bestFit").Text("M1", "no customWidth");
            tab.Text("N1", "default").Text("A2", "widest digit of the Normal font is the unit");

            tab.Text("A3", "J and K are both hidden; only K remembers how wide it was")
               .Expect(note: "hidden with a width and hidden with width 0 are different files and " +
                             "should be different documents");
            tab.Text("A4", "column 14 to 16384 share one <col> element")
               .Expect(note: "a span to the sheet's last column is a claim, not 16371 columns to build");
        }));
    }

    static void RowHeights(Corpus corpus)
    {
        var spec = New("geometry/rows.xlsx", "X4", "row heights in points, hidden rows, and zero height");
        spec.Note("Row heights are in points and need no conversion — which is what makes them the " +
                  "easy half, and worth stating because the column half looks the same and is not.");
        spec.Note("A row with `hidden=\"1\"` and a row with `ht=\"0\"` both disappear, and they are " +
                  "not the same thing: unhiding the first restores its height and the second has " +
                  "none to restore.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Rows");

            tab.FormatProperties = new SheetFormatProperties { DefaultRowHeight = 15 };

            var rows = new (int Index, double? Height, bool Hidden, string Label, string? Note)[]
            {
                (1, null, false, "default height", "no `ht` at all: the sheet's defaultRowHeight applies"),
                (2, 15, false, "15pt, stated", "the same height, written out — customHeight says the author chose it"),
                (3, 7.5, false, "7.5pt", "half the default"),
                (4, 30, false, "30pt", null),
                (5, 120.75, false, "120.75pt", "fractional points are legal"),
                (6, null, true, "hidden, no height", null),
                (7, 25, true, "hidden, remembers 25pt", "unhiding this restores 25pt"),
                (8, 0, false, "ht=0, not hidden", "zero height: invisible, and not the same as hidden"),
                (9, 409, false, "409pt", "Excel's maximum row height"),
            };

            foreach ((int index, double? height, bool hidden, string label, string? note) in rows)
            {
                tab.Row(index, height, hidden);
                tab.Text($"A{index}", label);
                if (note is not null) tab.Expect(note: note);
            }

            uint bold = book.Xf(new CellFormat
            {
                FontId = book.FontId(new Font(new Bold())),
                ApplyFont = true,
            });
            tab.Row(10, style: bold);
            tab.Text("A10", "row-level style")
               .Expect(note: "`customFormat` with an `s` on the row: a style for every cell in it, " +
                             "including the ones that do not exist");
        }));
    }

    static void Outlines(Corpus corpus)
    {
        var spec = New("geometry/outlines.xlsx", "X4", "grouped rows and columns, collapsed and not");
        spec.Note("Outline levels are Excel's row and column grouping. The model has hidden tracks " +
                  "but no grouping, so the *level* is the part with nowhere to go — a collapsed " +
                  "group's rows are hidden as well, and that half does carry.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Outlines");

            tab.Columns = new Columns(
                new Column { Min = 1, Max = 1, Width = 20, CustomWidth = true },
                new Column { Min = 2, Max = 4, Width = 10, OutlineLevel = 1, CustomWidth = true },
                new Column { Min = 5, Max = 5, Width = 10, OutlineLevel = 2, Hidden = true, CustomWidth = true });

            tab.Text("A1", "level 0").Text("B1", "level 1").Text("C1", "level 1")
               .Text("D1", "level 1").Text("E1", "level 2, hidden");

            tab.Row(2, outline: 0);
            tab.Text("A2", "ungrouped row");

            for (int row = 3; row <= 5; row++)
            {
                tab.Row(row, outline: 1);
                tab.Text($"A{row}", "group level 1");
            }

            for (int row = 6; row <= 7; row++)
            {
                tab.Row(row, outline: 2, hidden: true);
                tab.Text($"A{row}", "group level 2, collapsed");
            }

            tab.Row(8, collapsed: true);
            tab.Text("A8", "the summary row that owns the collapse")
               .Expect(note: "`collapsed` marks the row holding the group's toggle, not a hidden row");
        }));
    }

    static void Panes(Corpus corpus)
    {
        var spec = New("geometry/panes.xlsx", "X4", "frozen and split panes, and the selection");
        spec.Note("Panes are a view property rather than a document one. There is nowhere obvious " +
                  "for them in the model, and this file exists so that whatever happens to them is " +
                  "a decision with a fixture behind it rather than an omission.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab frozen = book.Sheet("FrozenTop");
            frozen.Views = new SheetViews(new SheetView(
                new Pane
                {
                    VerticalSplit = 1,
                    TopLeftCell = "A2",
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen,
                },
                new Selection { Pane = PaneValues.BottomLeft, ActiveCell = "A2", SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2" } })
            { WorkbookViewId = 0, TabSelected = true });

            for (int row = 1; row <= 5; row++)
                frozen.Text($"A{row}", row == 1 ? "header" : $"row {row}").Number($"B{row}", row);

            Tab both = book.Sheet("FrozenBoth");
            both.Views = new SheetViews(new SheetView(
                new Pane
                {
                    HorizontalSplit = 2,
                    VerticalSplit = 3,
                    TopLeftCell = "C4",
                    ActivePane = PaneValues.BottomRight,
                    State = PaneStateValues.Frozen,
                })
            { WorkbookViewId = 0 });
            both.Text("A1", "two frozen columns and three frozen rows");

            Tab split = book.Sheet("Split");
            split.Views = new SheetViews(new SheetView(
                new Pane
                {
                    HorizontalSplit = 2000,
                    VerticalSplit = 1000,
                    TopLeftCell = "D6",
                    ActivePane = PaneValues.BottomRight,
                    State = PaneStateValues.Split,
                })
            { WorkbookViewId = 0 });
            split.Text("A1", "a split, not a freeze — the offsets are twentieths of a point");

            Tab hidden = book.Sheet("NoGridlines");
            hidden.Views = new SheetViews(new SheetView
            {
                WorkbookViewId = 0,
                ShowGridLines = false,
                ShowRowColHeaders = false,
                ZoomScale = 150,
            });
            hidden.Text("A1", "gridlines and headers off, zoomed to 150%");
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "geometry", Milestone = milestone, Covers = covers };
}
