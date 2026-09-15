using OoxmlGen;
using OoxmlGen.Families;

namespace OoxmlGen;

/// <summary>
/// Generates a corpus of Excel workbooks for testing an importer against.
///
/// Every fixture is small, named for the construct it exercises, and paired with an entry in
/// <c>manifest.json</c> stating both what the file contains and what a conversion of it
/// should produce. Files are never hand-edited: this program is the source.
/// </summary>
public static class Program
{
    public const string Version = "0.2.0";

    /// <summary>
    /// The families, in the order they are generated. The names are what <c>--only</c> and
    /// <c>--skip</c> take.
    /// </summary>
    static readonly (string Name, string Summary, Action<Corpus> Run)[] Families =
    [
        ("values", "X1 — cell types, numbers, strings, the two date systems", Values.Generate),
        ("formulas", "X2 — expressions, shared groups, and the excluded classes", Formulas.Generate),
        ("numfmt", "X3 — built-in ids, custom codes, sections, currency, elapsed time", NumberFormats.Generate),
        ("styles", "X4 — fonts, colours, fills, borders, alignment, named styles", CellStyles.Generate),
        ("geometry", "X4 — column widths, row heights, hidden and outlined tracks", Geometry.Generate),
        ("document", "X5 — names, sheets, merges, filters, and the parts that get dropped", DocumentLevel.Generate),
        ("realworld", "X0 — what real producers write that the spec permits and nobody expects", RealWorld.Generate),
        ("hostile", "X0 — files that are trying to break the reader", Hostile.Generate),
        ("scale", "X1 — one large workbook, for the timing test", Scale.Generate),
    ];

    public static int Main(string[] args)
    {
        string output = Path.Combine(Directory.GetCurrentDirectory(), "corpus");
        var only = new List<string>();
        var skip = new List<string>();
        bool clean = true;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--out" or "-o" when i + 1 < args.Length:
                    output = Path.GetFullPath(args[++i]);
                    break;
                case "--only" when i + 1 < args.Length:
                    only.AddRange(args[++i].Split(',', StringSplitOptions.RemoveEmptyEntries));
                    break;
                case "--skip" when i + 1 < args.Length:
                    skip.AddRange(args[++i].Split(',', StringSplitOptions.RemoveEmptyEntries));
                    break;
                case "--no-clean":
                    clean = false;
                    break;
                case "--list":
                    foreach ((string name, string summary, _) in Families)
                        Console.WriteLine($"  {name,-11} {summary}");
                    return 0;
                case "--help" or "-h":
                    Usage();
                    return 0;
                default:
                    Console.Error.WriteLine($"ooxmlgen: unknown argument '{args[i]}'");
                    Usage();
                    return 2;
            }
        }

        foreach (string name in only.Concat(skip))
        {
            if (!Array.Exists(Families, f => f.Name == name))
            {
                Console.Error.WriteLine($"ooxmlgen: no such family '{name}' — try --list");
                return 2;
            }
        }

        var selected = Families
            .Where(f => (only.Count == 0 || only.Contains(f.Name)) && !skip.Contains(f.Name))
            .ToArray();

        if (clean && Directory.Exists(output)) Directory.Delete(output, recursive: true);
        Directory.CreateDirectory(output);

        var corpus = new Corpus(output);
        foreach ((string name, _, Action<Corpus> run) in selected)
        {
            Console.WriteLine($"{name}/");
            run(corpus);
        }

        corpus.WriteManifest(Version);

        long bytes = corpus.Fixtures.Sum(f => f.Bytes);
        int cells = corpus.Fixtures.Sum(f => f.Sheets.Sum(s => s.Cells.Count));
        Console.WriteLine();
        Console.WriteLine($"{corpus.Fixtures.Count} fixtures, {cells} asserted cells, " +
                          $"{bytes / 1024.0 / 1024.0:0.#} MiB → {output}");
        Console.WriteLine($"manifest: {Path.Combine(output, "manifest.json")}");
        return 0;
    }

    static void Usage()
    {
        Console.WriteLine("""
            ooxmlgen — generate a corpus of Excel workbooks for testing an importer

            usage: dotnet run -- [options]

              -o, --out <dir>     where to write the corpus (default: ./corpus)
                  --only <a,b>    generate only these families
                  --skip <a,b>    generate everything but these families
                  --no-clean      keep whatever is already in the output directory
                  --list          list the families and what each covers
              -h, --help          this

            The output directory gets one subdirectory per family, a manifest.json stating what
            every file contains and what a conversion of it should produce, and a README.md
            listing the lot.
            """);
    }
}
