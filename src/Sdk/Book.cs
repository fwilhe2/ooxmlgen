using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OoxmlGen.Sdk;

/// <summary>
/// A workbook under construction, plus the fixture's expectations about it.
///
/// Two things this takes off every family's hands. Cells and rows have to reach a worksheet
/// part in ascending order, and the children of <c>worksheet</c>, <c>workbook</c> and
/// <c>styleSheet</c> have to appear in the schema's sequence — both are silent corruption
/// when they are wrong, so the whole book is collected unordered and assembled in
/// <see cref="Finish"/>.
/// </summary>
public sealed class Book : IDisposable
{
    readonly SpreadsheetDocument doc;
    readonly WorkbookPart workbook;
    readonly FixtureSpec spec;
    readonly List<Tab> tabs = [];

    readonly List<SharedStringItem> strings = [];
    readonly Dictionary<string, int> plainStrings = [];

    readonly List<Font> fonts = [new Font()];
    readonly List<Fill> fills =
    [
        // Index 0 and 1 are fixed by convention: Excel writes them into every workbook it
        // creates and indexes user fills from 2, so a file that omits them is a file no
        // other producer looks like.
        new Fill(new PatternFill { PatternType = PatternValues.None }),
        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
    ];
    readonly List<Border> borders = [new Border()];
    readonly List<CellFormat> cellFormats = [new CellFormat { FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }];
    readonly List<CellFormat> cellStyleFormats = [new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }];
    readonly List<CellStyle> namedStyles = [];
    readonly List<NumberingFormat> numberFormats = [];

    uint nextNumberFormatId = 164;
    bool finished;

    Book(SpreadsheetDocument doc, FixtureSpec spec)
    {
        this.doc = doc;
        this.spec = spec;
        workbook = doc.AddWorkbookPart();
        workbook.Workbook = new Workbook();
    }

    /// <summary>Build a workbook at <paramref name="path"/>, recording into <paramref name="spec"/>.</summary>
    public static void Create(string path, FixtureSpec spec, SpreadsheetDocumentType type, Action<Book> build)
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, type))
        {
            using var book = new Book(document, spec);
            build(book);
            book.Finish();
        }

        Normalise(path);
    }

    /// <summary>
    /// Repack so that generating the corpus twice produces byte-identical files.
    ///
    /// Two things the SDK settles at random. It stamps every zip entry with the time it was
    /// written, and it names relationships with sixteen hex digits from a fresh GUID. Both
    /// make each run's hashes different, which turns every re-vendoring into a diff where
    /// nothing has actually changed and makes the manifest's hashes useless for pinning.
    ///
    /// The timestamp is fixed by writing through the corpus's own zip writer. The
    /// relationship ids are renumbered to rId1, rId2, … in a stable order, which is both
    /// deterministic and what every real producer writes anyway.
    /// </summary>
    static void Normalise(string path)
    {
        List<(string Name, byte[] Content)> parts = [];
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using Stream source = entry.Open();
                var buffer = new MemoryStream();
                source.CopyTo(buffer);
                parts.Add((entry.FullName, buffer.ToArray()));
            }
        }

        // Collected in name order rather than zip order, so the numbering does not depend on
        // the order parts happened to be added in.
        Dictionary<string, string> renamed = [];
        foreach ((string name, byte[] content) in parts.OrderBy(p => p.Name, StringComparer.Ordinal))
        {
            if (!IsText(name)) continue;
            foreach (Match match in GeneratedId.Matches(Text(content)))
                renamed.TryAdd(match.Value, $"rId{renamed.Count + 1}");
        }

        var zip = new Raw.ZipWriter();
        foreach ((string name, byte[] content) in parts)
        {
            zip.Add(name, IsText(name) ? Retext(content, renamed) : content);
        }

        zip.Save(path);
    }

    /// <summary>The SDK's generated relationship ids: an `R` and sixteen hex digits.</summary>
    static readonly System.Text.RegularExpressions.Regex GeneratedId = new("R[0-9a-f]{16}");

    static bool IsText(string name) =>
        name.EndsWith(".xml", StringComparison.Ordinal) || name.EndsWith(".rels", StringComparison.Ordinal);

    static string Text(byte[] content) =>
        System.Text.Encoding.UTF8.GetString(HasBom(content) ? content.AsSpan(3) : content);

    static bool HasBom(byte[] content) =>
        content.Length >= 3 && content[0] == 0xEF && content[1] == 0xBB && content[2] == 0xBF;

    static byte[] Retext(byte[] content, Dictionary<string, string> renamed)
    {
        string rewritten = GeneratedId.Replace(Text(content),
            match => renamed.TryGetValue(match.Value, out string? id) ? id : match.Value);
        byte[] body = System.Text.Encoding.UTF8.GetBytes(rewritten);
        return HasBom(content) ? [0xEF, 0xBB, 0xBF, .. body] : body;
    }

    public static void Create(string path, FixtureSpec spec, Action<Book> build) =>
        Create(path, spec, SpreadsheetDocumentType.Workbook, build);

    public WorkbookPart Part => workbook;

    public FixtureSpec Spec => spec;

    /// <summary>The 1904 date system. Off by default, as it is in every file Excel writes on Windows.</summary>
    public bool Date1904 { get; set; }

    public List<DefinedName> DefinedNames { get; } = [];

    public List<ExternalReference> ExternalReferences { get; } = [];

    public WorkbookProtection? Protection { get; set; }

    /// <summary>
    /// Children appended after <c>calcPr</c>, for the corners of CT_Workbook's sequence that
    /// have no typed slot here — <c>pivotCaches</c> is the one that needs it.
    /// </summary>
    public List<OpenXmlElement> ExtraWorkbookChildren { get; } = [];

    /// <summary>
    /// Write a part's XML directly. Used for the vocabularies that exist in this corpus only
    /// to be recognised and counted — DrawingML charts, pivot caches, VML — where building
    /// them through the object model would be a lot of code to express "and here is a chart".
    /// </summary>
    public static void WriteXml(OpenXmlPart part, string xml)
    {
        using Stream stream = part.GetStream(FileMode.Create);
        stream.Write(System.Text.Encoding.UTF8.GetBytes(xml));
    }

    // --- sheets ---

    /// <summary>Add a sheet. The order they are added in is the order the workbook lists them.</summary>
    public Tab Sheet(string name, SheetStateValues? state = null, string? expectName = null)
    {
        WorksheetPart part = workbook.AddNewPart<WorksheetPart>();
        var sheetSpec = new SheetSpec
        {
            Name = name,
            ExpectName = expectName,
            Hidden = state is not null,
        };
        spec.Sheets.Add(sheetSpec);

        var tab = new Tab(this, part, sheetSpec, state);
        tabs.Add(tab);
        return tab;
    }

    // --- shared strings ---

    /// <summary>Intern a plain string, returning its index in the shared string table.</summary>
    public int String(string text)
    {
        if (plainStrings.TryGetValue(text, out int existing)) return existing;

        var item = new SharedStringItem(new Text(text));
        // A string whose ends are whitespace loses them unless it says not to. Silent text
        // corruption, and the reason `xml:space` is in the real-world family at all.
        if (text.Length != text.Trim().Length) item.Text!.Space = SpaceProcessingModeValues.Preserve;

        strings.Add(item);
        int index = strings.Count - 1;
        plainStrings[text] = index;
        return index;
    }

    /// <summary>Intern a multi-run (rich text) string. Never deduplicated — each is its own item.</summary>
    public int RichString(params (string Text, RunProperties? Style)[] runs)
    {
        var item = new SharedStringItem();
        foreach ((string text, RunProperties? style) in runs)
        {
            var run = new Run();
            if (style is not null) run.Append(style);
            var element = new Text(text);
            if (text.Length != text.Trim().Length) element.Space = SpaceProcessingModeValues.Preserve;
            run.Append(element);
            item.Append(run);
        }

        strings.Add(item);
        return strings.Count - 1;
    }

    // --- styles ---

    /// <summary>
    /// Attach a theme part, which is what gives <c>&lt;color theme="4" tint="-0.25"/&gt;</c>
    /// anything to resolve against.
    ///
    /// The index a cell writes is *not* the order the colours appear in the scheme: index 0
    /// and 1 are lt1 and dk1, in that order, and so are 2 and 3 for lt2/dk2 — the light and
    /// dark of each pair are swapped relative to the XML. Indices 4 to 9 are accent1 to
    /// accent6, 10 is hlink and 11 is folHlink. The scheme below gives each of the twelve a
    /// distinguishable colour so a fixture can tell a correct mapping from a plausible one.
    /// </summary>
    public void AddTheme()
    {
        ThemePart part = workbook.AddNewPart<ThemePart>();
        using Stream stream = part.GetStream(FileMode.Create);
        stream.Write(System.Text.Encoding.UTF8.GetBytes(ThemeXml));
    }

    /// <summary>The colours <see cref="AddTheme"/> installs, by the index a cell would use.</summary>
    public static readonly string[] ThemeColours =
    [
        "FFFFFF", // 0  lt1
        "000000", // 1  dk1
        "E7E6E6", // 2  lt2
        "44546A", // 3  dk2
        "4472C4", // 4  accent1
        "ED7D31", // 5  accent2
        "A5A5A5", // 6  accent3
        "FFC000", // 7  accent4
        "5B9BD5", // 8  accent5
        "70AD47", // 9  accent6
        "0563C1", // 10 hlink
        "954F72", // 11 folHlink
    ];

    const string ThemeXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
          <a:themeElements>
            <a:clrScheme name="Office">
              <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
              <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
              <a:dk2><a:srgbClr val="44546A"/></a:dk2>
              <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
              <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
              <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
              <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
              <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
              <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
              <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
              <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
              <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="Office">
              <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
              <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
            </a:fontScheme>
            <a:fmtScheme name="Office">
              <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
              </a:fillStyleLst>
              <a:lnStyleLst>
                <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
              </a:lnStyleLst>
              <a:effectStyleLst>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
              </a:effectStyleLst>
              <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
              </a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>
        """;

    public uint FontId(Font font)
    {
        fonts.Add(font);
        return (uint)(fonts.Count - 1);
    }

    public uint FillId(Fill fill)
    {
        fills.Add(fill);
        return (uint)(fills.Count - 1);
    }

    public uint BorderId(Border border)
    {
        borders.Add(border);
        return (uint)(borders.Count - 1);
    }

    /// <summary>Define a custom number format, returning its id. Built-ins keep their own 0–163.</summary>
    public uint NumberFormatId(string code)
    {
        NumberingFormat? existing = numberFormats.Find(f => f.FormatCode?.Value == code);
        if (existing is not null) return existing.NumberFormatId!.Value;

        uint id = nextNumberFormatId++;
        numberFormats.Add(new NumberingFormat { NumberFormatId = id, FormatCode = code });
        return id;
    }

    /// <summary>Append a cell format, returning the <c>s</c> index that selects it.</summary>
    public uint Xf(CellFormat format)
    {
        format.FormatId ??= 0;
        cellFormats.Add(format);
        return (uint)(cellFormats.Count - 1);
    }

    /// <summary>A cell format whose only job is to apply a number format.</summary>
    public uint Format(string code) =>
        Xf(new CellFormat { NumberFormatId = NumberFormatId(code), ApplyNumberFormat = true });

    /// <summary>A cell format selecting one of ECMA-376 §18.8.30's built-in format ids.</summary>
    public uint BuiltinFormat(uint id) =>
        Xf(new CellFormat { NumberFormatId = id, ApplyNumberFormat = true });

    /// <summary>Add a named cell style, which is a <c>cellStyleXfs</c> entry plus a name for it.</summary>
    public uint NamedStyle(string name, CellFormat format, uint builtinId)
    {
        cellStyleFormats.Add(format);
        uint xfId = (uint)(cellStyleFormats.Count - 1);
        namedStyles.Add(new CellStyle { Name = name, FormatId = xfId, BuiltinId = builtinId });
        return xfId;
    }

    // --- assembly ---

    void Finish()
    {
        if (finished) return;
        finished = true;

        WorkbookStylesPart stylesPart = workbook.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet();
        stylesPart.Stylesheet.Save();

        if (strings.Count > 0)
        {
            SharedStringTablePart part = workbook.AddNewPart<SharedStringTablePart>();
            var table = new SharedStringTable
            {
                Count = (uint)strings.Count,
                UniqueCount = (uint)strings.Count,
            };
            foreach (SharedStringItem item in strings) table.Append(item);
            part.SharedStringTable = table;
            part.SharedStringTable.Save();
        }

        var sheets = new Sheets();
        uint sheetId = 1;
        foreach (Tab tab in tabs)
        {
            tab.Assemble();
            var sheet = new Sheet
            {
                Id = workbook.GetIdOfPart(tab.PartOf),
                SheetId = sheetId++,
                Name = tab.SheetName,
            };
            if (tab.State is { } state) sheet.State = state;
            sheets.Append(sheet);
        }

        // CT_Workbook's sequence: workbookPr, workbookProtection, bookViews, sheets,
        // externalReferences, definedNames, calcPr.
        Workbook book = workbook.Workbook;
        if (Date1904) book.Append(new WorkbookProperties { Date1904 = true });
        if (Protection is not null) book.Append(Protection);
        book.Append(sheets);
        if (ExternalReferences.Count > 0)
        {
            var references = new ExternalReferences();
            foreach (ExternalReference reference in ExternalReferences) references.Append(reference);
            book.Append(references);
        }

        if (DefinedNames.Count > 0)
        {
            var names = new DefinedNames();
            foreach (DefinedName name in DefinedNames) names.Append(name);
            book.Append(names);
        }

        // Cached values are the whole point of this corpus: nothing may invite a consumer to
        // throw them away and recompute on open.
        book.Append(new CalculationProperties { CalculationId = 191029, FullCalculationOnLoad = false });
        foreach (OpenXmlElement child in ExtraWorkbookChildren) book.Append(child);
        book.Save();
    }

    Stylesheet BuildStylesheet()
    {
        // CT_Stylesheet's sequence: numFmts, fonts, fills, borders, cellStyleXfs, cellXfs,
        // cellStyles.
        var stylesheet = new Stylesheet();

        if (numberFormats.Count > 0)
        {
            var formats = new NumberingFormats { Count = (uint)numberFormats.Count };
            foreach (NumberingFormat format in numberFormats) formats.Append(format);
            stylesheet.Append(formats);
        }

        var fontList = new Fonts { Count = (uint)fonts.Count };
        foreach (Font font in fonts) fontList.Append(font);
        stylesheet.Append(fontList);

        var fillList = new Fills { Count = (uint)fills.Count };
        foreach (Fill fill in fills) fillList.Append(fill);
        stylesheet.Append(fillList);

        var borderList = new Borders { Count = (uint)borders.Count };
        foreach (Border border in borders) borderList.Append(border);
        stylesheet.Append(borderList);

        var styleFormats = new CellStyleFormats { Count = (uint)cellStyleFormats.Count };
        foreach (CellFormat format in cellStyleFormats) styleFormats.Append(format);
        stylesheet.Append(styleFormats);

        var formatList = new CellFormats { Count = (uint)cellFormats.Count };
        foreach (CellFormat format in cellFormats) formatList.Append(format);
        stylesheet.Append(formatList);

        if (namedStyles.Count > 0)
        {
            var styles = new CellStyles { Count = (uint)namedStyles.Count };
            foreach (CellStyle style in namedStyles) styles.Append(style);
            stylesheet.Append(styles);
        }
        else
        {
            stylesheet.Append(new CellStyles(
                new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });
        }

        return stylesheet;
    }

    public void Dispose()
    {
        Finish();
        doc.Dispose();
    }
}
