using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OoxmlGen.Sdk;

/// <summary>
/// One sheet. Every method that writes a cell also records what a conversion of that cell
/// should produce, which is the only reason the manifest can be trusted: the two halves are
/// never edited apart.
/// </summary>
public sealed class Tab
{
    readonly Book book;
    readonly SheetSpec spec;
    readonly SortedDictionary<int, RowBuild> rows = [];
    CellSpec? last;

    internal Tab(Book book, WorksheetPart part, SheetSpec spec, SheetStateValues? state)
    {
        this.book = book;
        this.spec = spec;
        PartOf = part;
        State = state;
    }

    internal WorksheetPart PartOf { get; }

    internal SheetStateValues? State { get; }

    internal string SheetName => spec.Name;

    // Slots for the children of `worksheet`, which has to emit them in schema order. Holding
    // them separately is what lets a family set them in any order it likes.
    public SheetProperties? Properties { get; set; }
    public SheetDimension? Dimension { get; set; }
    public SheetViews? Views { get; set; }
    public SheetFormatProperties? FormatProperties { get; set; }
    public Columns? Columns { get; set; }
    public SheetProtection? Protection { get; set; }
    public AutoFilter? AutoFilter { get; set; }
    public MergeCells? Merges { get; set; }
    public List<ConditionalFormatting> ConditionalFormats { get; } = [];
    public DataValidations? Validations { get; set; }
    public Drawing? Drawing { get; set; }
    public LegacyDrawing? LegacyDrawing { get; set; }
    public TableParts? TableParts { get; set; }

    // --- rows ---

    /// <summary>Set a row's own attributes — height, visibility, outline level.</summary>
    public Tab Row(int index, double? height = null, bool hidden = false, byte? outline = null,
                   bool collapsed = false, uint? style = null)
    {
        Row element = At(index).Element;
        if (height is { } h)
        {
            element.Height = h;
            element.CustomHeight = true;
        }

        if (hidden) element.Hidden = true;
        if (outline is { } level) element.OutlineLevel = level;
        if (collapsed) element.Collapsed = true;
        if (style is { } s)
        {
            element.StyleIndex = s;
            element.CustomFormat = true;
        }

        return this;
    }

    /// <summary>Drop the <c>r</c> attribute from a row, leaving its position implicit.</summary>
    public Tab ImplicitRow(int index)
    {
        At(index).Implicit = true;
        return this;
    }

    // --- cells ---

    /// <summary>A number, written the way Excel writes one: no <c>t</c> at all.</summary>
    public Tab Number(string reference, double value, uint style = 0) =>
        Number(reference, value.ToString("R", CultureInfo.InvariantCulture), style);

    /// <summary>A number whose literal spelling in <c>&lt;v&gt;</c> matters.</summary>
    public Tab Number(string reference, string raw, uint style = 0) =>
        Put(reference, null, raw, style, ValueKind.Number, raw);

    /// <summary>Text through the shared string table, which is how Excel stores nearly all of it.</summary>
    public Tab Text(string reference, string text, uint style = 0) =>
        Put(reference, CellValues.SharedString, book.String(text).ToString(), style, ValueKind.Text, text);

    /// <summary>Text held in the cell itself. Common from non-Excel producers.</summary>
    public Tab Inline(string reference, string text, uint style = 0)
    {
        var element = new Text(text);
        if (text.Length != text.Trim().Length) element.Space = SpaceProcessingModeValues.Preserve;
        Tab tab = Put(reference, CellValues.InlineString, null, style, ValueKind.Text, text,
                      inline: new InlineString(element));
        last!.Raw = text;
        return tab;
    }

    /// <summary>A multi-run shared string. More than one run has nowhere to go in ODF's cell model.</summary>
    public Tab Rich(string reference, string flattened, params (string Text, RunProperties? Style)[] runs) =>
        Put(reference, CellValues.SharedString, book.RichString(runs).ToString(), 0, ValueKind.Text, flattened);

    public Tab Bool(string reference, bool value, uint style = 0) =>
        Put(reference, CellValues.Boolean, value ? "1" : "0", style, ValueKind.Bool, value ? "true" : "false");

    /// <summary>An error constant. <paramref name="name"/> is Excel's spelling, which is also ODF's.</summary>
    public Tab Error(string reference, string name, uint style = 0) =>
        Put(reference, CellValues.Error, name, style, ValueKind.Error, name);

    /// <summary>ECMA-376 2nd edition's <c>t="d"</c>: an ISO 8601 date in the cell, no serial anywhere.</summary>
    public Tab IsoDate(string reference, string iso, uint style = 0) =>
        Put(reference, CellValues.Date, iso, style, ValueKind.Date, iso);

    /// <summary>A serial that is a date only because its format says so.</summary>
    public Tab Serial(string reference, double serial, uint style, string? iso, string? display = null)
    {
        Number(reference, serial, style);
        return Expect(kind: ValueKind.Date, value: iso, display: display);
    }

    /// <summary>An empty but styled cell: <c>&lt;c r="A1" s="3"/&gt;</c>, which real files are full of.</summary>
    public Tab Blank(string reference, uint style = 0) =>
        Put(reference, null, null, style, ValueKind.Empty, null);

    /// <summary>
    /// A formula and the value Excel cached for it. <paramref name="expect"/> is the
    /// OpenFormula spelling a translation should arrive at, or null where the fixture
    /// asserts the formula is not carried at all.
    /// </summary>
    public Tab Formula(string reference, string excel, string? cached, string? expect,
                       CellValues? type = null, uint style = 0, ValueKind kind = ValueKind.Number)
    {
        string? value = type is null || type.Value == CellValues.Number ? cached : null;
        Put(reference, type, cached, style, kind, value ?? cached, formula: new CellFormula(excel));
        last!.Excel = excel;
        last.Formula = expect;
        return this;
    }

    /// <summary>The master of a shared-formula group: it carries the expression and the range.</summary>
    public Tab SharedFormulaMaster(string reference, uint index, string range, string excel,
                                   string? cached, string? expect, uint style = 0)
    {
        var formula = new CellFormula(excel)
        {
            FormulaType = CellFormulaValues.Shared,
            SharedIndex = index,
            Reference = range,
        };
        Put(reference, null, cached, style, ValueKind.Number, cached, formula: formula);
        last!.Excel = excel;
        last.Formula = expect;
        return this;
    }

    /// <summary>
    /// A follower in a shared-formula group: an empty <c>&lt;f&gt;</c> that means "the master's
    /// expression, with its relative references shifted to here".
    /// </summary>
    public Tab SharedFormulaMember(string reference, uint index, string? cached, string? expect, uint style = 0)
    {
        var formula = new CellFormula
        {
            FormulaType = CellFormulaValues.Shared,
            SharedIndex = index,
        };
        Put(reference, null, cached, style, ValueKind.Number, cached, formula: formula);
        last!.Formula = expect;
        last.Note ??= $"shared group si={index}; the expression is the master's, shifted";
        return this;
    }

    /// <summary>An array formula. Out of scope by grind's §2.3.2: the value stays, the formula goes.</summary>
    public Tab ArrayFormula(string reference, string range, string excel, string? cached, uint style = 0)
    {
        var formula = new CellFormula(excel)
        {
            FormulaType = CellFormulaValues.Array,
            Reference = range,
        };
        Put(reference, null, cached, style, ValueKind.Number, cached, formula: formula);
        last!.Excel = excel;
        last.Formula = null;
        last.Note ??= "array formula — the cached value is carried, the formula is dropped";
        return this;
    }

    /// <summary>Drop the <c>r</c> attribute from the cell most recently written.</summary>
    public Tab Implicit()
    {
        Last.Note ??= "no `r` attribute — the position is implicit";
        implicitCells.Add(Last.Ref);
        return this;
    }

    // --- expectations ---

    /// <summary>Adjust what the most recently written cell is expected to convert to.</summary>
    public Tab Expect(ValueKind? kind = null, string? value = null, string? display = null,
                      string? formula = null, string? note = null)
    {
        CellSpec cell = Last;
        if (kind is { } k) cell.Kind = k;
        if (value is not null) cell.Value = value;
        if (display is not null) cell.Display = display;
        if (formula is not null) cell.Formula = formula;
        if (note is not null) cell.Note = note;
        return this;
    }

    /// <summary>State that the most recently written cell converts to nothing assertable.</summary>
    public Tab ExpectNoValue(string note)
    {
        Last.Value = null;
        Last.Note = note;
        return this;
    }

    /// <summary>
    /// Record what LibreOffice's own conversion produces for the most recent cell, where that
    /// has been measured and differs from what the file means.
    /// </summary>
    public Tab Oracle(string value)
    {
        Last.Oracle = value;
        return this;
    }

    /// <summary>
    /// A duration, spelled the way ODF spells one. A number under a time format is a duration
    /// rather than a clock reading: the format decides how it reads, not what it is.
    /// </summary>
    public static string Duration(double days)
    {
        double seconds = Math.Round(days * 86400.0, 6);
        long hours = (long)(seconds / 3600);
        long minutes = (long)(seconds % 3600 / 60);
        long rest = (long)(seconds % 60);
        return $"PT{hours:00}H{minutes:00}M{rest:00}S";
    }

    CellSpec Last => last ?? throw new InvalidOperationException("no cell has been written yet");

    // --- plumbing ---

    readonly HashSet<string> implicitCells = [];

    Tab Put(string reference, CellValues? type, string? raw, uint style, ValueKind kind,
            string? value, CellFormula? formula = null, InlineString? inline = null)
    {
        (int col, int row) = A1.Parse(reference);

        var cell = new Cell { CellReference = reference };
        if (type is { } t) cell.DataType = t;
        if (style != 0) cell.StyleIndex = style;

        // CT_Cell's sequence is f, v, is — and a cached value after its formula is the
        // whole point of this corpus.
        if (formula is not null) cell.Append(formula);
        if (raw is not null && inline is null) cell.Append(new CellValue(raw));
        if (inline is not null) cell.Append(inline);

        At(row).Cells[col] = cell;

        last = new CellSpec
        {
            Ref = reference,
            T = WireType(type),
            Raw = raw,
            Style = style == 0 ? null : style,
            Kind = kind,
            Value = value,
        };
        spec.Cells.Add(last);
        return this;
    }

    /// <summary>
    /// The `t` attribute as it appears in the file. In this SDK's version `CellValues` is a
    /// struct whose `ToString` gives the type's name rather than the wire value, so the
    /// mapping is written out — and the manifest records what the bytes say rather than what
    /// the object model calls it.
    /// </summary>
    static string? WireType(CellValues? type) => type switch
    {
        null => null,
        var t when t.Value == CellValues.Number => "n",
        var t when t.Value == CellValues.SharedString => "s",
        var t when t.Value == CellValues.String => "str",
        var t when t.Value == CellValues.InlineString => "inlineStr",
        var t when t.Value == CellValues.Boolean => "b",
        var t when t.Value == CellValues.Error => "e",
        var t when t.Value == CellValues.Date => "d",
        var t => t.Value.ToString(),
    };

    RowBuild At(int index)
    {
        if (!rows.TryGetValue(index, out RowBuild? build))
        {
            build = new RowBuild();
            rows[index] = build;
        }

        return build;
    }

    /// <summary>Emit the worksheet, in ascending order and in the schema's sequence.</summary>
    internal void Assemble()
    {
        var data = new SheetData();
        foreach ((int index, RowBuild build) in rows)
        {
            Row row = build.Element;
            if (!build.Implicit) row.RowIndex = (uint)index;
            foreach (Cell cell in build.Cells.Values)
            {
                if (implicitCells.Contains(cell.CellReference!.Value!)) cell.CellReference = null;
                row.Append(cell);
            }

            data.Append(row);
        }

        var sheet = new Worksheet();
        if (Properties is not null) sheet.Append(Properties);
        if (Dimension is not null) sheet.Append(Dimension);
        if (Views is not null) sheet.Append(Views);
        if (FormatProperties is not null) sheet.Append(FormatProperties);
        if (Columns is not null) sheet.Append(Columns);
        sheet.Append(data);
        if (Protection is not null) sheet.Append(Protection);
        if (AutoFilter is not null) sheet.Append(AutoFilter);
        if (Merges is not null) sheet.Append(Merges);
        foreach (ConditionalFormatting format in ConditionalFormats) sheet.Append(format);
        if (Validations is not null) sheet.Append(Validations);
        if (Drawing is not null) sheet.Append(Drawing);
        if (LegacyDrawing is not null) sheet.Append(LegacyDrawing);
        if (TableParts is not null) sheet.Append(TableParts);

        PartOf.Worksheet = sheet;
        PartOf.Worksheet.Save();
    }

    sealed class RowBuild
    {
        public Row Element { get; } = new();
        public SortedDictionary<int, Cell> Cells { get; } = [];
        public bool Implicit { get; set; }
    }
}
