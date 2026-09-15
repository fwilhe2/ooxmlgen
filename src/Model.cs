using System.Text.Json.Serialization;

namespace OoxmlGen;

/// <summary>
/// The kind of value an importer is expected to end up with, in ODF's vocabulary rather than
/// Excel's. Excel has one numeric type and lets the number format decide what it means; ODF
/// carries the distinction on the cell, so this is what a conversion has to arrive at.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<ValueKind>))]
public enum ValueKind
{
    Empty,
    Number,
    Text,
    Bool,
    Error,
    Date,
    Time,
    Percentage,
    Currency,
}

/// <summary>
/// Which of ECMA-376's two namespace families the file uses. A fact about the file, not a
/// mode a reader runs in.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<Flavour>))]
public enum Flavour
{
    Transitional,
    Strict,
    Mixed,
}

/// <summary>
/// Constructs an importer is expected to drop and count. The names are grind's
/// <c>Dropped</c> enum verbatim so a manifest entry can be matched against a report without
/// a translation table in between.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<Dropped>))]
public enum Dropped
{
    Chart,
    PivotTable,
    ConditionalFormat,
    DataValidation,
    Comment,
    Drawing,
    Macro,
    ArrayFormula,
    StructuredReference,
    ExternalLink,
    SheetLocalName,
    MergedCells,
    RichText,
    HiddenSheet,
    ThemeColor,
    FontFamily,
    Protection,
}

/// <summary>
/// How a file is expected to fail outright. Only the two cases that are a whole-file
/// verdict: everything else is a dropped construct, because a conversion that refuses a
/// workbook over one bad chart is a conversion nobody can use.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<ExpectedError>))]
public enum ExpectedError
{
    /// <summary>A CFB/OLE container: the package is there but locked.</summary>
    Encrypted,

    /// <summary>Not a readable OPC package at all.</summary>
    Package,
}

/// <summary>
/// One cell, stated twice: what the file literally contains, and what a conversion of it
/// should produce. Both halves are written at the point the cell is generated, so the
/// manifest cannot drift away from the bytes.
/// </summary>
public sealed class CellSpec
{
    /// <summary>The A1 reference. Present even when the file omits <c>r</c>.</summary>
    public required string Ref { get; init; }

    /// <summary>The <c>t</c> attribute as written, or <c>null</c> where the file omits it.</summary>
    public string? T { get; set; }

    /// <summary>The literal content of <c>&lt;v&gt;</c>, exactly as it appears in the part.</summary>
    public string? Raw { get; set; }

    /// <summary>The <c>s</c> attribute, where the cell carries a style.</summary>
    public uint? Style { get; set; }

    /// <summary>The value kind a conversion should arrive at.</summary>
    public ValueKind Kind { get; set; }

    /// <summary>
    /// The expected value in canonical form: a number as an invariant decimal, a date or
    /// time as ISO 8601, a bool as <c>true</c>/<c>false</c>, an error as its name, text as
    /// itself. Null where the fixture deliberately asserts nothing.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>The text the cell should display once its number format is applied.</summary>
    public string? Display { get; set; }

    /// <summary>The Excel expression as written in <c>&lt;f&gt;</c>, without a leading <c>=</c>.</summary>
    public string? Excel { get; set; }

    /// <summary>The expected translation, in OpenFormula's canonical spelling.</summary>
    public string? Formula { get; set; }

    /// <summary>
    /// What <c>soffice --headless --convert-to ods</c> actually produces for this cell, where
    /// that has been measured and differs from <see cref="Value"/>. A conversion compared
    /// against LibreOffice as an oracle will disagree here by the oracle's fault, not its
    /// own, and a test that does not know which cells those are reads as a failure.
    /// </summary>
    public string? Oracle { get; set; }

    /// <summary>Why this cell is interesting, where that is not obvious from the values.</summary>
    public string? Note { get; set; }
}

/// <summary>One sheet of a fixture, and everything asserted about it.</summary>
public sealed class SheetSpec
{
    public required string Name { get; init; }

    /// <summary>The name a conversion should end up with, where ODF cannot spell Excel's.</summary>
    public string? ExpectName { get; set; }

    public bool Hidden { get; set; }

    public List<CellSpec> Cells { get; } = [];
}

/// <summary>One generated file, and the whole of what it is for.</summary>
public sealed class FixtureSpec
{
    /// <summary>Path relative to the corpus root, with forward slashes.</summary>
    public required string File { get; init; }

    public required string Family { get; init; }

    /// <summary>The milestone in grind's <c>doc/xlsx-import.md</c> Part III this exercises.</summary>
    public required string Milestone { get; init; }

    /// <summary>One line on the construct under test.</summary>
    public required string Covers { get; init; }

    public Flavour Flavour { get; set; } = Flavour.Transitional;

    /// <summary>Set only where the whole file is expected to fail to open.</summary>
    public ExpectedError? Error { get; set; }

    /// <summary>
    /// Whether <c>soffice --headless --convert-to ods</c> opens this file, measured rather
    /// than assumed. It is not a statement about what an importer *should* do — several files
    /// here the oracle opens and a stricter reader may rightly refuse, and one it refuses that
    /// a better reader could handle. It is here so a test harness knows which conversions to
    /// expect output from, without hard-coding a list of file names.
    /// </summary>
    public bool OracleOpens { get; set; } = true;

    /// <summary>
    /// Whether the package deliberately stores two entries under one name. A structural fact
    /// a checker needs to know, because everywhere else a duplicate entry is a generator bug.
    /// </summary>
    public bool DuplicateEntries { get; set; }

    public List<SheetSpec> Sheets { get; } = [];

    /// <summary>Constructs that should appear in the report, by kind and count.</summary>
    public SortedDictionary<Dropped, int> ExpectDropped { get; } = [];

    /// <summary>Functions a carried formula names that a Small-Group build will not know.</summary>
    public SortedSet<string> ExpectUnknownFunctions { get; } = [];

    /// <summary>Namespaces the file asserts a consumer must understand.</summary>
    public SortedSet<string> ExpectMustUnderstand { get; } = [];

    /// <summary>Anything a reader should know that the fields above cannot say.</summary>
    public List<string> Notes { get; } = [];

    /// <summary>Size on disk, filled in once the file exists.</summary>
    public long Bytes { get; set; }

    /// <summary>Content hash, so a vendored copy can be pinned.</summary>
    public string? Sha256 { get; set; }

    /// <summary>Record a construct the conversion is expected to drop.</summary>
    public FixtureSpec Drops(Dropped kind, int count = 1)
    {
        ExpectDropped.TryGetValue(kind, out int seen);
        ExpectDropped[kind] = seen + count;
        return this;
    }

    public FixtureSpec Unknown(params string[] functions)
    {
        foreach (string f in functions) ExpectUnknownFunctions.Add(f);
        return this;
    }

    public FixtureSpec Note(string note)
    {
        Notes.Add(note);
        return this;
    }
}

/// <summary>The corpus as a whole, as written to <c>manifest.json</c>.</summary>
public sealed class ManifestSpec
{
    public required string Generator { get; init; }
    public required string Schema { get; init; }
    public required string Generated { get; init; }
    public required List<FixtureSpec> Fixtures { get; init; }
}
