using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// Every function grind implements, called on realistic data — odsgen's function samples
/// (<c>generate_function_samples.py</c>), written again as Excel workbooks.
///
/// odsgen is the source: it holds the data and the formulas once, in OpenFormula, has
/// LibreOffice recalculate them, and writes <c>functions/samples.json</c> with every cell as
/// LibreOffice saved it and each formula's Excel spelling. This family reads that file
/// (<c>--function-samples</c>, by default <c>../odsgen/functions/samples.json</c>) and writes
/// one workbook per category with the same sheets, the same filter and the same formulas.
///
/// The Open XML SDK does not calculate, so each formula's cached <c>&lt;v&gt;</c> is
/// LibreOffice's answer, and the manifest says so: the claims are that the import carries the
/// value and translates the formula back to the OpenFormula it came from. Whether a
/// recalculation reproduces the value is the consumer's own loop. A formula with no Excel
/// spelling — OpenFormula's three-parameter <c>CEILING</c> and <c>FLOOR</c> — is left out.
/// </summary>
public static class Functions
{
    /// <summary>Where odsgen's samples are; set from the command line.</summary>
    public static string SamplesPath { get; set; } =
        Path.Combine("..", "odsgen", "functions", "samples.json");

    public static void Generate(Corpus corpus)
    {
        if (!File.Exists(SamplesPath))
            throw new FileNotFoundException(
                $"the functions family reads odsgen's samples, and {SamplesPath} is not there — " +
                "run odsgen's generate_function_samples.py, or pass --function-samples", SamplesPath);

        using JsonDocument json = JsonDocument.Parse(File.ReadAllText(SamplesPath));
        foreach (JsonElement document in json.RootElement.GetProperty("documents").EnumerateArray())
            Workbook(corpus, document);
    }

    static void Workbook(Corpus corpus, JsonElement document)
    {
        string slug = document.GetProperty("slug").GetString()!;
        string title = document.GetProperty("title").GetString()!;
        var spec = new FixtureSpec
        {
            File = $"functions/{slug}.xlsx",
            Family = "functions",
            Milestone = "X2",
            Covers = $"{title.ToLowerInvariant()} functions on realistic data: " +
                     document.GetProperty("story").GetString(),
        };
        spec.Note("Generated from odsgen's function samples. The cached values are LibreOffice's " +
                  "answers to the same formulas in OpenFormula, not Excel's: the Open XML SDK " +
                  "does not calculate. `formula` is the OpenFormula the sample was written in.");
        spec.Note("Samples whose formula Excel cannot spell — OpenFormula's own CEILING and FLOOR, " +
                  "whose Mode parameter Excel's namesakes lack — are left out.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint date = book.Format("yyyy-mm-dd");
            uint time = book.Format("hh:mm:ss");
            foreach (JsonElement sheet in document.GetProperty("sheets").EnumerateArray())
                Sheet(book, sheet, date, time);
        }));
    }

    static void Sheet(Book book, JsonElement sheet, uint date, uint time)
    {
        string name = sheet.GetProperty("name").GetString()!;
        Tab tab = book.Sheet(name);
        bool samples = name == "Samples";
        var skipped = new HashSet<int>();

        foreach (JsonElement cell in sheet.GetProperty("cells").EnumerateArray())
        {
            string reference = cell.GetProperty("ref").GetString()!;
            int row = A1.Parse(reference).Row;
            if (skipped.Contains(row)) continue;

            if (cell.TryGetProperty("formula", out _))
            {
                string? excel = cell.GetProperty("excel").GetString();
                if (excel is null)
                {
                    // A sample Excel cannot say: leave its whole row out of the workbook.
                    skipped.Add(row);
                    continue;
                }
                Formula(tab, reference, excel, cell, date, time);
            }
            else if (samples && reference.StartsWith('B') && row > 1)
            {
                // The formula as typed — in Excel's spelling here, as Excel's formula bar shows it.
                string? excel = sheet.GetProperty("cells").EnumerateArray()
                    .First(c => c.GetProperty("ref").GetString() == $"C{row}")
                    .GetProperty("excel").GetString();
                if (excel is null)
                {
                    skipped.Add(row);
                    continue;
                }
                tab.Text(reference, "=" + excel);
            }
            else
            {
                Value(tab, reference, cell, date, time);
            }
        }

        if (sheet.TryGetProperty("hidden", out JsonElement hidden))
            foreach (JsonProperty row in hidden.EnumerateObject())
                tab.Row(int.Parse(row.Name, CultureInfo.InvariantCulture), hidden: true);

        if (sheet.TryGetProperty("filter", out JsonElement filter) && filter.ValueKind == JsonValueKind.Array)
        {
            int lastRow = filter[0].GetInt32();
            uint field = (uint)filter[1].GetInt32();
            // The filter covers the table's own width: as far as its heading row goes.
            int last = sheet.GetProperty("cells").EnumerateArray()
                .Select(c => A1.Parse(c.GetProperty("ref").GetString()!))
                .Where(p => p.Row == 1)
                .Max(p => p.Col);
            var values = new Filters();
            foreach (JsonElement value in filter[2].EnumerateArray())
                values.Append(new Filter { Val = value.GetString() });
            tab.AutoFilter = new AutoFilter(new FilterColumn(values) { ColumnId = field })
            {
                Reference = $"A1:{A1.Column(last)}{lastRow}",
            };
        }
    }

    /// <summary>A plain cell, written the way Excel writes its kind.</summary>
    static void Value(Tab tab, string reference, JsonElement cell, uint date, uint time)
    {
        string? kind = cell.GetProperty("kind").GetString();
        string? value = cell.GetProperty("value").GetString();
        switch (kind)
        {
            case "float" or "percentage" or "currency":
                tab.Number(reference, double.Parse(value!, CultureInfo.InvariantCulture));
                break;
            case "boolean":
                tab.Bool(reference, value == "true");
                break;
            case "date":
                tab.Serial(reference, Serial(value!), date, null);
                tab.Expect(value: Serial(value!).ToString("R", CultureInfo.InvariantCulture));
                break;
            case "time":
                double fraction = Fraction(value!);
                tab.Serial(reference, fraction, time, null);
                tab.Expect(kind: ValueKind.Time, value: fraction.ToString("R", CultureInfo.InvariantCulture));
                break;
            default:
                tab.Text(reference, value ?? "");
                break;
        }
    }

    /// <summary>A formula, with LibreOffice's answer as its cached value.</summary>
    static void Formula(Tab tab, string reference, string excel, JsonElement cell, uint date, uint time)
    {
        string odf = cell.GetProperty("odf").GetString()!;
        string? kind = cell.GetProperty("kind").GetString();
        string? value = cell.GetProperty("value").GetString();
        switch (kind)
        {
            case "float" or "percentage" or "currency":
                tab.Formula(reference, excel, value, odf);
                break;
            case "boolean":
                tab.Formula(reference, excel, value == "true" ? "1" : "0", odf, CellValues.Boolean, kind: ValueKind.Bool)
                   .Expect(value: value);
                break;
            case "date":
            {
                string serial = Serial(value!).ToString("R", CultureInfo.InvariantCulture);
                tab.Formula(reference, excel, serial, odf, style: date).Expect(kind: ValueKind.Date, value: serial);
                break;
            }
            case "time":
            {
                string fraction = Fraction(value!).ToString("R", CultureInfo.InvariantCulture);
                tab.Formula(reference, excel, fraction, odf, style: time).Expect(kind: ValueKind.Time, value: fraction);
                break;
            }
            case "error":
                tab.Formula(reference, excel, value, odf, CellValues.Error, kind: ValueKind.Error);
                break;
            default:
                // Text — including the empty text `T` gives a number, which LibreOffice saves as
                // a cell with no value at all.
                tab.Formula(reference, excel, value ?? "", odf, CellValues.String, kind: ValueKind.Text);
                break;
        }
    }

    /// <summary>An ISO 8601 date (or date-time) as a 1900-system serial; every one here is past 1900-03-01.</summary>
    static double Serial(string iso)
    {
        DateTime when = DateTime.Parse(iso, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        return (when - new DateTime(1899, 12, 30)).TotalDays;
    }

    /// <summary>An ODF duration, <c>PT08H47M30S</c>, as a fraction of a day.</summary>
    static double Fraction(string duration) =>
        System.Xml.XmlConvert.ToTimeSpan(duration).TotalDays;
}
