using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace OoxmlGen.Raw;

/// <summary>
/// A zip writer with no opinions.
///
/// <see cref="System.IO.Compression.ZipArchive"/> is the right tool for writing an archive a
/// reader should accept, and the wrong one for writing the archives in this corpus: it
/// normalises entry names, picks its own compression, and orders the central directory to
/// suit itself. Half the real-world and hostile fixtures exist precisely to exercise a name,
/// an order, or a compression method a well-behaved writer would never produce.
/// </summary>
public sealed class ZipWriter
{
    readonly List<Entry> entries = [];

    /// <summary>Add a part from bytes already in hand, replacing any entry of the same name.</summary>
    public void Add(string name, byte[] content, bool deflate = true)
    {
        byte[] stored = deflate ? Deflate(content) : content;
        Put(new Entry(name, stored, Crc32.Of(content), content.Length, deflate));
    }

    /// <summary>
    /// Install an entry, replacing one already under that name.
    ///
    /// Two entries sharing a name makes a file whose meaning depends on which one the reader
    /// happens to reach — and a fixture that lands on the wrong one tests nothing. Replacing
    /// is what a caller overriding an assembled part always means. A fixture that genuinely
    /// wants a duplicate can have <see cref="AddDuplicate"/>.
    /// </summary>
    void Put(Entry entry)
    {
        int existing = entries.FindIndex(e => e.Name == entry.Name);
        if (existing >= 0) entries[existing] = entry;
        else entries.Add(entry);
    }

    /// <summary>Add a second entry under a name already taken, deliberately.</summary>
    public void AddDuplicate(string name, string content)
    {
        byte[] body = Encoding.UTF8.GetBytes(content);
        entries.Add(new Entry(name, Deflate(body), Crc32.Of(body), body.Length, Deflated: true));
    }

    /// <summary>Add a part from text, encoded UTF-8, optionally with a byte order mark.</summary>
    public void Add(string name, string content, bool bom = false, bool deflate = true)
    {
        byte[] body = Encoding.UTF8.GetBytes(content);
        if (bom) body = [.. Preamble, .. body];
        Add(name, body, deflate);
    }

    /// <summary>
    /// Add a part whose uncompressed form is too large to hold in memory. The producer writes
    /// into a stream that computes the CRC as the bytes pass through it and deflates them on
    /// the way out, so only the compressed result is ever materialised.
    /// </summary>
    public void AddStreamed(string name, Action<Stream> produce)
    {
        var compressed = new MemoryStream();
        var crc = new Crc32();
        long length;

        using (var deflate = new DeflateStream(compressed, CompressionLevel.SmallestSize, leaveOpen: true))
        using (var counting = new CountingCrcStream(deflate, crc))
        {
            produce(counting);
            counting.Flush();
            length = counting.Written;
        }

        Put(new Entry(name, compressed.ToArray(), crc.Value, length, Deflated: true));
    }

    /// <summary>Write the archive.</summary>
    public void Save(string path)
    {
        using FileStream file = File.Create(path);
        Write(file);
    }

    /// <summary>
    /// Write only the first <paramref name="fraction"/> of the archive — a file that starts
    /// out looking exactly like a workbook and stops in the middle of one.
    /// </summary>
    public void SaveTruncated(string path, double fraction)
    {
        var whole = new MemoryStream();
        Write(whole);
        byte[] bytes = whole.ToArray();
        File.WriteAllBytes(path, bytes[..(int)(bytes.Length * fraction)]);
    }

    void Write(Stream output)
    {
        var offsets = new List<long>(entries.Count);

        foreach (Entry entry in entries)
        {
            offsets.Add(output.Position);
            byte[] name = Encoding.UTF8.GetBytes(entry.Name);

            Span<byte> header = stackalloc byte[30];
            BinaryPrimitives.WriteUInt32LittleEndian(header[0..], 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header[4..], 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header[6..], 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[8..], entry.Deflated ? (ushort)8 : (ushort)0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[10..], DosTime);
            BinaryPrimitives.WriteUInt16LittleEndian(header[12..], DosDate);
            BinaryPrimitives.WriteUInt32LittleEndian(header[14..], entry.Crc);
            BinaryPrimitives.WriteUInt32LittleEndian(header[18..], (uint)entry.Content.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(header[22..], (uint)entry.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header[26..], (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0);

            output.Write(header);
            output.Write(name);
            output.Write(entry.Content);
        }

        long directory = output.Position;
        for (int i = 0; i < entries.Count; i++)
        {
            Entry entry = entries[i];
            byte[] name = Encoding.UTF8.GetBytes(entry.Name);

            Span<byte> header = stackalloc byte[46];
            BinaryPrimitives.WriteUInt32LittleEndian(header[0..], 0x02014b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header[4..], 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header[6..], 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header[8..], 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[10..], entry.Deflated ? (ushort)8 : (ushort)0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[12..], DosTime);
            BinaryPrimitives.WriteUInt16LittleEndian(header[14..], DosDate);
            BinaryPrimitives.WriteUInt32LittleEndian(header[16..], entry.Crc);
            BinaryPrimitives.WriteUInt32LittleEndian(header[20..], (uint)entry.Content.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(header[24..], (uint)entry.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header[28..], (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[34..], 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header[36..], 0);
            BinaryPrimitives.WriteUInt32LittleEndian(header[38..], 0);
            BinaryPrimitives.WriteUInt32LittleEndian(header[42..], (uint)offsets[i]);

            output.Write(header);
            output.Write(name);
        }

        long directorySize = output.Position - directory;

        Span<byte> end = stackalloc byte[22];
        BinaryPrimitives.WriteUInt32LittleEndian(end[0..], 0x06054b50);
        BinaryPrimitives.WriteUInt16LittleEndian(end[4..], 0);
        BinaryPrimitives.WriteUInt16LittleEndian(end[6..], 0);
        BinaryPrimitives.WriteUInt16LittleEndian(end[8..], (ushort)entries.Count);
        BinaryPrimitives.WriteUInt16LittleEndian(end[10..], (ushort)entries.Count);
        BinaryPrimitives.WriteUInt32LittleEndian(end[12..], (uint)directorySize);
        BinaryPrimitives.WriteUInt32LittleEndian(end[16..], (uint)directory);
        BinaryPrimitives.WriteUInt16LittleEndian(end[20..], 0);
        output.Write(end);
    }

    static byte[] Deflate(byte[] content)
    {
        var output = new MemoryStream();
        using (var stream = new DeflateStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
            stream.Write(content);
        return output.ToArray();
    }

    /// <summary>A fixed timestamp, so the same corpus generates byte-identical twice running.</summary>
    const ushort DosTime = 0;
    const ushort DosDate = (2020 - 1980) << 9 | 1 << 5 | 1;

    static ReadOnlySpan<byte> Preamble => [0xEF, 0xBB, 0xBF];

    readonly record struct Entry(string Name, byte[] Content, uint Crc, long Length, bool Deflated);

    /// <summary>Counts and checksums what passes through, then forwards it.</summary>
    sealed class CountingCrcStream(Stream inner, Crc32 crc) : Stream
    {
        public long Written { get; private set; }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            crc.Append(buffer);
            Written += buffer.Length;
            inner.Write(buffer);
        }

        public override void Write(byte[] buffer, int offset, int count) =>
            Write(buffer.AsSpan(offset, count));

        public override void Flush() => inner.Flush();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => Written;
        public override long Position { get => Written; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}

/// <summary>The zip checksum, incrementally.</summary>
public sealed class Crc32
{
    static readonly uint[] Table = Build();

    uint state = 0xFFFFFFFF;

    public uint Value => state ^ 0xFFFFFFFF;

    public void Append(ReadOnlySpan<byte> bytes)
    {
        uint value = state;
        foreach (byte b in bytes) value = Table[(value ^ b) & 0xFF] ^ value >> 8;
        state = value;
    }

    public static uint Of(ReadOnlySpan<byte> bytes)
    {
        var crc = new Crc32();
        crc.Append(bytes);
        return crc.Value;
    }

    static uint[] Build()
    {
        var table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint value = i;
            for (int bit = 0; bit < 8; bit++)
                value = (value & 1) != 0 ? 0xEDB88320 ^ value >> 1 : value >> 1;
            table[i] = value;
        }

        return table;
    }
}
