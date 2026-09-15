namespace OoxmlGen.Raw;

/// <summary>
/// A workbook assembled from XML text rather than through an object model.
///
/// Everything a well-behaved producer settles for you — the namespace family, the byte order
/// mark, the XML declaration, how a relationship spells its target, whether a part is
/// deflated at all — is a knob here, because those are exactly the decisions the real-world
/// and hostile families need to make differently.
/// </summary>
public sealed class RawWorkbook
{
    // ECMA-376's two namespace families. Parts 2 and 3 — the package, content types and
    // markup compatibility — do not vary between them, which is half of why Strict costs so
    // little to recognise.
    public const string Transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string Strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    public const string TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string StrictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    public const string MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    public const string TransitionalRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string StrictRelType = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    const string SheetContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const string WorkbookContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    const string StylesContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    const string SharedStringsContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

    /// <summary>Use the ISO Strict namespace family throughout.</summary>
    public bool UseStrict { get; set; }

    /// <summary>Put a UTF-8 byte order mark in front of every part.</summary>
    public bool Bom { get; set; }

    /// <summary>The XML declaration each part opens with, or null for none at all.</summary>
    public string? Declaration { get; set; } = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""";

    /// <summary>The zip entry the workbook part is actually stored under.</summary>
    public string WorkbookPath { get; set; } = "xl/workbook.xml";

    /// <summary>How the root relationship spells its target. Not required to be a clean path.</summary>
    public string? WorkbookTarget { get; set; }

    /// <summary>The zip entry each worksheet is stored under, in sheet order.</summary>
    public List<string> SheetPaths { get; } = ["xl/worksheets/sheet1.xml"];

    /// <summary>How the workbook relationship spells each sheet's target, in sheet order.</summary>
    public List<string>? SheetTargets { get; set; }

    /// <summary>Sheet names, in sheet order.</summary>
    public List<string> SheetNames { get; } = ["Sheet1"];

    /// <summary>The body of each worksheet — the `sheetData` and anything around it.</summary>
    public List<string> SheetBodies { get; } = [];

    /// <summary>Extra attributes on the `worksheet` root, such as an `mc:Ignorable` list.</summary>
    public string RootAttributes { get; set; } = "";

    /// <summary>Extra namespace declarations on the `worksheet` root.</summary>
    public List<string> ExtraNamespaces { get; } = [];

    /// <summary>Attributes on the `workbook` root.</summary>
    public string WorkbookAttributes { get; set; } = "";

    /// <summary>The shared string table, if the workbook has one.</summary>
    public string? SharedStrings { get; set; }

    /// <summary>Parts written verbatim, over and above the ones assembled here.</summary>
    public List<(string Name, string Xml)> ExtraParts { get; } = [];

    /// <summary>Content-type overrides for <see cref="ExtraParts"/>.</summary>
    public List<(string PartName, string ContentType)> ExtraOverrides { get; } = [];

    /// <summary>Entries stored rather than deflated.</summary>
    public HashSet<string> Stored { get; } = [];

    /// <summary>Applied to the writer just before it is saved, for anything else.</summary>
    public Action<ZipWriter>? Finally { get; set; }

    public string Main => UseStrict ? Strict : Transitional;

    public string Relationships => UseStrict ? StrictRelationships : TransitionalRelationships;

    string RelType => UseStrict ? StrictRelType : TransitionalRelType;

    public void Save(string path)
    {
        var zip = new ZipWriter();

        string workbookTarget = WorkbookTarget ?? "/" + WorkbookPath;
        var sheetTargets = SheetTargets ?? [.. SheetPaths.Select(Relative)];

        zip.Add("[Content_Types].xml", Part(ContentTypesXml()), Bom, !Stored.Contains("[Content_Types].xml"));
        zip.Add("_rels/.rels", Part($"""
            <Relationships xmlns="{PackageRelationships}">
              <Relationship Id="rId1" Type="{RelType}/officeDocument" Target="{workbookTarget}"/>
            </Relationships>
            """), Bom);

        zip.Add(WorkbookPath, Part(WorkbookXml()), Bom, !Stored.Contains(WorkbookPath));
        zip.Add(RelsFor(WorkbookPath), Part(WorkbookRelsXml(sheetTargets)), Bom);

        for (int i = 0; i < SheetPaths.Count; i++)
        {
            // A sheet supplied verbatim through ExtraParts replaces the assembled one rather
            // than joining it: two zip entries under one name is a file whose meaning depends
            // on which one the reader happens to reach.
            if (ExtraParts.Any(p => p.Name == SheetPaths[i])) continue;

            string body = i < SheetBodies.Count ? SheetBodies[i] : "<sheetData/>";
            zip.Add(SheetPaths[i], Part(SheetXml(body)), Bom, !Stored.Contains(SheetPaths[i]));
        }

        zip.Add("xl/styles.xml", Part(StylesXml()), Bom);

        if (SharedStrings is not null)
            zip.Add("xl/sharedStrings.xml", Part(SharedStrings), Bom);

        foreach ((string name, string xml) in ExtraParts)
            zip.Add(name, Part(xml), Bom, !Stored.Contains(name));

        Finally?.Invoke(zip);
        zip.Save(path);
    }

    string Part(string body) => Declaration is null ? body.TrimStart() : Declaration + "\n" + body.TrimStart();

    static string RelsFor(string part)
    {
        int slash = part.LastIndexOf('/');
        return slash < 0 ? $"_rels/{part}.rels" : $"{part[..slash]}/_rels/{part[(slash + 1)..]}.rels";
    }

    /// <summary>A sheet path expressed relative to the workbook part's own directory.</summary>
    string Relative(string sheetPath)
    {
        string directory = WorkbookPath.Contains('/') ? WorkbookPath[..WorkbookPath.LastIndexOf('/')] + "/" : "";
        return sheetPath.StartsWith(directory, StringComparison.Ordinal)
            ? sheetPath[directory.Length..]
            : "/" + sheetPath;
    }

    string ContentTypesXml()
    {
        var overrides = new List<string>
        {
            $"""<Override PartName="/{WorkbookPath}" ContentType="{WorkbookContentType}"/>""",
            $"""<Override PartName="/xl/styles.xml" ContentType="{StylesContentType}"/>""",
        };

        overrides.AddRange(SheetPaths.Select(p =>
            $"""<Override PartName="/{p}" ContentType="{SheetContentType}"/>"""));

        if (SharedStrings is not null)
            overrides.Add($"""<Override PartName="/xl/sharedStrings.xml" ContentType="{SharedStringsContentType}"/>""");

        overrides.AddRange(ExtraOverrides.Select(o =>
            $"""<Override PartName="{o.PartName}" ContentType="{o.ContentType}"/>"""));

        return $"""
            <Types xmlns="{ContentTypesNs}">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              {string.Join("\n  ", overrides)}
            </Types>
            """;
    }

    string WorkbookXml()
    {
        string sheets = string.Join("\n    ", SheetNames.Select((name, i) =>
            $"""<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>"""));

        return $"""
            <workbook xmlns="{Main}" xmlns:r="{Relationships}"{WorkbookAttributes}>
              <sheets>
                {sheets}
              </sheets>
            </workbook>
            """;
    }

    string WorkbookRelsXml(List<string> sheetTargets)
    {
        var relationships = sheetTargets.Select((target, i) =>
            $"""<Relationship Id="rId{i + 1}" Type="{RelType}/worksheet" Target="{target}"/>""").ToList();

        relationships.Add(
            $"""<Relationship Id="rIdStyles" Type="{RelType}/styles" Target="styles.xml"/>""");

        if (SharedStrings is not null)
            relationships.Add(
                $"""<Relationship Id="rIdStrings" Type="{RelType}/sharedStrings" Target="sharedStrings.xml"/>""");

        return $"""
            <Relationships xmlns="{PackageRelationships}">
              {string.Join("\n  ", relationships)}
            </Relationships>
            """;
    }

    string SheetXml(string body)
    {
        string extra = ExtraNamespaces.Count == 0 ? "" : " " + string.Join(" ", ExtraNamespaces);
        return $"""
            <worksheet xmlns="{Main}" xmlns:r="{Relationships}"{extra}{RootAttributes}>
              {body}
            </worksheet>
            """;
    }

    string StylesXml() => $"""
        <styleSheet xmlns="{Main}">
          <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\-mm\-dd"/></numFmts>
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;
}
