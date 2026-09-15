using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X1 — cell values and the two date systems. The half of an import that everything else is
/// measured against: a formula's cached value only means something once the values around it
/// are known to be right.
/// </summary>
public static class Values
{
    public static void Generate(Corpus corpus)
    {
        Types(corpus);
        Numbers(corpus);
        Strings(corpus);
        RichText(corpus);
        InlineStrings(corpus);
        EmptyAndSparse(corpus);
        Dates1900(corpus);
        Dates1904(corpus);
        DateVersusNumber(corpus);
        Times(corpus);
    }

    /// <summary>Every value of the <c>t</c> attribute, one per row.</summary>
    static void Types(Corpus corpus)
    {
        var spec = New("values/types.xlsx", "X1", "every cell type: n, s, str, inlineStr, b, e, d");
        spec.Note("`t=\"d\"` is ECMA-376 2nd edition. Excel does not write it; other producers do.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint date = book.Format("yyyy-mm-dd");
            Tab tab = book.Sheet("Types");

            tab.Text("A1", "t").Text("B1", "value");

            tab.Text("A2", "n (implicit)").Number("B2", 42.5);
            tab.Text("A3", "s").Text("B3", "shared string");
            tab.Text("A4", "str").Formula("B4", "\"formula\"&\" result\"", "formula result",
                                          "\"formula\"&\" result\"", CellValues.String, kind: ValueKind.Text);
            tab.Text("A5", "inlineStr").Inline("B5", "inline string");
            tab.Text("A6", "b true").Bool("B6", true);
            tab.Text("A7", "b false").Bool("B7", false);
            tab.Text("A8", "d").IsoDate("B8", "2024-03-17", date);

            // ECMA-376 §18.18.11's whole error set. It is also OpenFormula §5.12's, which is
            // why these carry across unrenamed.
            int row = 9;
            foreach (string error in new[] { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" })
            {
                tab.Text($"A{row}", $"e {error}").Error($"B{row}", error);
                row++;
            }
        }));
    }

    /// <summary>Numbers at the edges of what a double can say, and what a file can spell.</summary>
    static void Numbers(Corpus corpus)
    {
        var spec = New("values/numbers.xlsx", "X1", "numeric literals: precision, magnitude, spelling");
        spec.Note("LibreOffice writes 15 significant digits, so a comparison against the oracle " +
                  "cannot assert more than that.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Numbers");

            tab.Text("A1", "case").Text("B1", "value");

            var cases = new (string Label, string Raw, string? Note)[]
            {
                ("zero", "0", null),
                ("negative zero", "-0", "a distinction the file can make and the model need not keep"),
                ("integer", "1", null),
                ("negative", "-273.15", null),
                ("one third", "0.333333333333333", "15 significant digits — LibreOffice's limit"),
                ("pi to 17", "3.1415926535897931", "17 digits: a double's round-trip spelling"),
                ("scientific", "1.2345E-7", "`<v>` may hold exponent notation; Excel writes it for small values"),
                ("large", "1E+25", null),
                ("largest exact integer", "9007199254740992", "2^53 — past here, integers are lossy"),
                ("excel max", "1.7976931348623157E+308", null),
                ("excel min positive", "2.2250738585072014E-308", null),
                ("leading plus", "+5", "the spec permits it; almost nothing writes it"),
                ("trailing zeros", "1.500", "significant to the file, not to the value"),
            };

            int row = 2;
            foreach ((string label, string raw, string? note) in cases)
            {
                tab.Text($"A{row}", label).Number($"B{row}", raw);
                if (note is not null) tab.Expect(note: note);
                row++;
            }
        }));
    }

    /// <summary>The shared string table, and everything text can contain.</summary>
    static void Strings(Corpus corpus)
    {
        var spec = New("values/strings.xlsx", "X1", "shared strings: whitespace, Unicode, XML escapes, dedup");
        spec.Note("B6 and B7 hold the same text and must share one table entry: an importer that " +
                  "indexes by position rather than by content reads B7 as something else.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Strings");

            tab.Text("A1", "case").Text("B1", "value");

            var cases = new (string Label, string Text, string? Note)[]
            {
                ("empty", "", "an empty `<t/>`, not an empty cell"),
                ("plain", "hello", null),
                ("leading space", "  indented", "needs `xml:space=\"preserve\"` or it is corrupted silently"),
                ("trailing space", "trailing  ", "likewise"),
                ("only spaces", "   ", null),
                ("duplicate a", "the same text", "shares a table entry with the row below"),
                ("duplicate b", "the same text", null),
                ("xml escapes", "a & b < c > d \" e ' f", "every character XML has to escape"),
                ("cjk", "日本語のテキスト", null),
                ("rtl", "نص عربي", null),
                ("emoji", "spreadsheet 📊 and a flag 🇩🇪", "astral plane: surrogate pairs in UTF-16"),
                ("combining", "égal", "e + combining acute, not the precomposed é"),
                ("newline", "first line\nsecond line", "a literal newline inside `<t>`"),
                ("tab", "a\tb", null),
                ("long", new string('x', 4096), "well past what any inline buffer should assume"),
            };

            int row = 2;
            foreach ((string label, string text, string? note) in cases)
            {
                tab.Text($"A{row}", label).Text($"B{row}", text);
                if (note is not null) tab.Expect(note: note);
                row++;
            }
        }));
    }

    /// <summary>Multi-run strings: the one text construct ODF's cell model has no home for.</summary>
    static void RichText(Corpus corpus)
    {
        var spec = New("values/rich-text.xlsx", "X1", "multi-run shared strings flatten, and are counted");
        spec.Drops(Dropped.RichText, 3);
        spec.Note("A single-run item with an `rPr` is still rich text in the file and still flattens, " +
                  "but carries no information the flattened form loses — B2 is the case that argues " +
                  "the count should be of items that lose something, not of items that had an `rPr`.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("RichText");

            tab.Text("A1", "case").Text("B1", "value");

            tab.Text("A2", "one run, styled")
               .Rich("B2", "all bold", ("all bold", new RunProperties(new Bold())));

            tab.Text("A3", "two runs")
               .Rich("B3", "plain and bold",
                     ("plain and ", null),
                     ("bold", new RunProperties(new Bold())));

            tab.Text("A4", "many runs")
               .Rich("B4", "red big italic",
                     ("red ", new RunProperties(new Color { Rgb = "FFFF0000" })),
                     ("big ", new RunProperties(new FontSize { Val = 18 })),
                     ("italic", new RunProperties(new Italic())));

            tab.Text("A5", "runs with whitespace")
               .Rich("B5", "a  b",
                     ("a ", null),
                     (" b", null))
               .Expect(note: "the space belongs to neither run alone — flattening must not trim either");
        }));
    }

    /// <summary>A workbook with no shared string part at all.</summary>
    static void InlineStrings(Corpus corpus)
    {
        var spec = New("values/inline-strings.xlsx", "X1", "all text inline; no sharedStrings part exists");
        spec.Note("A reader that reaches for `xl/sharedStrings.xml` before checking whether the " +
                  "relationship exists fails on this file, and on everything SheetJS writes by default.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Inline");
            tab.Inline("A1", "product").Inline("B1", "quantity");
            tab.Inline("A2", "widget").Number("B2", 12);
            tab.Inline("A3", "gadget").Number("B3", 7);
            tab.Inline("A4", "  padded  ").Number("B4", 0)
               .Expect(note: "inline text needs `xml:space` exactly as a shared string does");
        }));
    }

    /// <summary>Holes: the shapes a sheet takes when most of it is nothing.</summary>
    static void EmptyAndSparse(Corpus corpus)
    {
        var spec = New("values/empty-and-sparse.xlsx", "X1", "self-closed cells, gaps, and an empty sheet");
        spec.Note("Sheet 'Empty' has a `sheetData` with no rows in it at all — valid, and a shape " +
                  "a reader that assumes at least one row will not survive.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint bold = book.Xf(new CellFormat { FontId = book.FontId(new Font(new Bold())), ApplyFont = true });

            Tab sparse = book.Sheet("Sparse");
            sparse.Text("A1", "top left");
            sparse.Blank("C1", bold).Expect(note: "styled but empty: it has an `s` and no `v`");
            sparse.Number("E1", 5);
            // Rows 2-9 do not exist at all.
            sparse.Number("A10", 10);
            sparse.Text("Z10", "far right");
            sparse.Blank("A11").Expect(note: "`<c r=\"A11\"/>` — no style, no value, still present");
            sparse.Number("AA100", 100).Expect(note: "past Z: the two-letter column spelling");

            book.Sheet("Empty");
        }));
    }

    /// <summary>The 1900 system, including the day in 1900 that never happened.</summary>
    static void Dates1900(Corpus corpus)
    {
        var spec = New("values/dates-1900.xlsx", "X1", "the 1900 epoch and the phantom 1900-02-29");
        spec.Note("Serial 60 is Excel's leap-year bug: 1900 was not a leap year, but Lotus 1-2-3 " +
                  "believed it was and Excel kept the belief for compatibility. ODF's epoch is " +
                  "1899-12-30, so serials 1–59 shift by one and serials 61 and up already agree.");
        spec.Note("The correction applies only because these cells carry a date format. The same " +
                  "serials with a numeric format are plain numbers — see date-vs-number.xlsx.");
        spec.Note("`oracle` on the cells below is LibreOffice's own answer, measured on 2026-09-14 " +
                  "with `soffice --headless --convert-to ods`. It applies no correction at all: it " +
                  "adds the serial to 1899-12-30 throughout, so every serial from 1 to 60 comes out " +
                  "a day earlier than Excel displays it. A comparison against the oracle will " +
                  "disagree on rows 2 to 5 of this file, and the oracle is the one that is wrong.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint iso = book.Format("yyyy-mm-dd");
            Tab tab = book.Sheet("Dates1900");

            tab.Text("A1", "serial").Text("B1", "date").Text("C1", "note");

            // The oracle's answer is simply 1899-12-30 + serial, which is what the two columns
            // disagree about below serial 61.
            var epoch = new DateTime(1899, 12, 30);
            var cases = new (double Serial, string? Iso, string? Note)[]
            {
                (1, "1900-01-01", "the epoch, as Excel displays it"),
                (2, "1900-01-02", "still below the phantom day"),
                (59, "1900-02-28", "the last real day before the bug"),
                (60, null, "1900-02-29 in Excel: a day that does not exist, and nothing to assert"),
                (61, "1900-03-01", "the first day Excel and ODF agree exactly"),
                (62, "1900-03-02", "past the bug — carried unchanged from here on"),
                (366, "1900-12-31", null),
                (367, "1901-01-01", null),
                (25569, "1970-01-01", "the Unix epoch, for anyone checking with a different tool"),
                (45368, "2024-03-17", "a date from this century"),
                (2958465, "9999-12-31", "Excel's last representable day"),
            };

            int row = 2;
            foreach ((double serial, string? isoDate, string? note) in cases)
            {
                tab.Number($"A{row}", serial);
                tab.Serial($"B{row}", serial, iso, isoDate, isoDate);
                if (note is not null) tab.Expect(note: note);
                if (isoDate is null) tab.ExpectNoValue(note!);

                string oracle = epoch.AddDays(serial).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                if (oracle != isoDate) tab.Oracle(oracle);

                if (note is not null) tab.Text($"C{row}", note);
                row++;
            }
        }));
    }

    /// <summary>The 1904 system: the same calendar dates, different serials, no bug.</summary>
    static void Dates1904(Corpus corpus)
    {
        var spec = New("values/dates-1904.xlsx", "X1", "the 1904 epoch — serials pass through unshifted");
        spec.Note("`workbookPr/@date1904` is the whole switch. ODF carries the epoch per document " +
                  "in `table:null-date`, so this is a translation rather than an arithmetic correction.");
        spec.Note("The dates here are deliberately the same calendar days as dates-1900.xlsx from " +
                  "1904 onwards, so the two files can be compared cell for cell.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            book.Date1904 = true;
            uint iso = book.Format("yyyy-mm-dd");
            Tab tab = book.Sheet("Dates1904");

            tab.Text("A1", "serial").Text("B1", "date");

            var epoch = new DateTime(1904, 1, 1);
            var cases = new (string Iso, string? Note)[]
            {
                ("1904-01-01", "serial 0 — the epoch itself"),
                ("1904-01-02", null),
                ("1904-02-29", "1904 really was a leap year; no phantom anywhere in this system"),
                ("1904-03-01", null),
                ("1970-01-01", null),
                ("2024-03-17", null),
                ("9999-12-31", null),
            };

            int row = 2;
            foreach ((string isoDate, string? note) in cases)
            {
                double serial = (DateTime.Parse(isoDate, CultureInfo.InvariantCulture) - epoch).TotalDays;
                tab.Number($"A{row}", serial);
                tab.Serial($"B{row}", serial, iso, isoDate, isoDate);
                if (note is not null) tab.Expect(note: note);
                row++;
            }
        }));
    }

    /// <summary>One serial, five formats: proof the format is what makes a number a date.</summary>
    static void DateVersusNumber(Corpus corpus)
    {
        var spec = New("values/date-vs-number.xlsx", "X1",
                       "the same serial as date, time, datetime and plain number");
        spec.Note("Every cell in column B holds the identical `<v>45368.75</v>`. Only `s` differs. " +
                  "A reader that applies the epoch correction without consulting the format shifts " +
                  "B6 by a day and reports 45369 for a cell that is not a date at all.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint date = book.Format("yyyy-mm-dd");
            uint time = book.Format("hh:mm:ss");
            uint stamp = book.Format("yyyy-mm-dd hh:mm:ss");
            uint plain = book.Format("0.00");
            uint general = 0;

            Tab tab = book.Sheet("SameSerial");
            tab.Text("A1", "format").Text("B1", "value");

            const string raw = "45368.75";

            tab.Text("A2", "yyyy-mm-dd").Number("B2", raw, date)
               .Expect(ValueKind.Date, "2024-03-17T18:00:00", "2024-03-17")
               .Expect(note: "a date-only format does not truncate the value — the time is still there, " +
                             "it is only not shown");
            tab.Text("A3", "hh:mm:ss").Number("B3", raw, time)
               .Expect(ValueKind.Time, Tab.Duration(45368.75), "18:00:00")
               .Expect(note: "a time format makes the whole serial a duration, not a clock reading: " +
                             "45368.75 days is PT1088850H. The format wraps it to 18:00 for display, " +
                             "which is a fact about the format rather than about the value");
            tab.Text("A4", "yyyy-mm-dd hh:mm:ss").Number("B4", raw, stamp)
               .Expect(ValueKind.Date, "2024-03-17T18:00:00", "2024-03-17 18:00:00");
            tab.Text("A5", "0.00").Number("B5", raw, plain)
               .Expect(ValueKind.Number, raw, "45368.75")
               .Expect(note: "a numeric format: no epoch correction, no date kind");
            tab.Text("A6", "General").Number("B6", raw, general)
               .Expect(ValueKind.Number, raw)
               .Expect(note: "no format at all is still not a date");
        }));
    }

    /// <summary>Times, including the ones past midnight that are not times of day.</summary>
    static void Times(Corpus corpus)
    {
        var spec = New("values/times.xlsx", "X1", "times of day, and durations past 24 hours");
        spec.Note("Columns B and C hold identical values and differ only in format. The converted " +
                  "*value* is therefore the same duration in both — measured: LibreOffice writes " +
                  "`PT36H00M00S` for 1.5 under either. It is the *display* that differs, because " +
                  "`hh:mm:ss` wraps at 24 hours and `[h]:mm:ss` accumulates. A fixture that asserted " +
                  "the value differed would be asserting the format changed the number.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint clock = book.Format("hh:mm:ss");
            uint elapsed = book.Format("[h]:mm:ss");
            Tab tab = book.Sheet("Times");

            tab.Text("A1", "serial").Text("B1", "as clock").Text("C1", "as elapsed");

            var cases = new (string Raw, string Clock, string Elapsed)[]
            {
                ("0", "00:00:00", "0:00:00"),
                ("0.25", "06:00:00", "6:00:00"),
                ("0.5", "12:00:00", "12:00:00"),
                ("0.99998842592592593", "23:59:59", "23:59:59"),
                ("1", "00:00:00", "24:00:00"),
                ("1.5", "12:00:00", "36:00:00"),
                ("2.75", "18:00:00", "66:00:00"),
            };

            int row = 2;
            foreach ((string raw, string clockText, string elapsedText) in cases)
            {
                string duration = Tab.Duration(double.Parse(raw, CultureInfo.InvariantCulture));
                tab.Text($"A{row}", raw);
                tab.Number($"B{row}", raw, clock).Expect(ValueKind.Time, duration, clockText)
                   .Expect(note: "a clock format wraps at 24 hours");
                tab.Number($"C{row}", raw, elapsed).Expect(ValueKind.Time, duration, elapsedText)
                   .Expect(note: "the same value; elapsed hours accumulate instead of wrapping");
                row++;
            }
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "values", Milestone = milestone, Covers = covers };
}
