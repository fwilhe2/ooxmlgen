namespace OoxmlGen;

/// <summary>
/// A1 references, both ways. Cells and rows have to reach a worksheet part in ascending
/// order or the file is invalid, so every fixture builds an unordered bag and sorts on the
/// way out — which is what these are for.
/// </summary>
public static class A1
{
    /// <summary>Split an A1 reference into its one-based column and row.</summary>
    public static (int Col, int Row) Parse(string reference)
    {
        int i = 0;
        int col = 0;
        while (i < reference.Length && char.IsAsciiLetter(reference[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
            i++;
        }

        if (col == 0 || i == reference.Length)
            throw new ArgumentException($"not an A1 reference: {reference}", nameof(reference));

        return (col, int.Parse(reference[i..]));
    }

    /// <summary>Spell a one-based column number the way Excel does: A, Z, AA, XFD.</summary>
    public static string Column(int col)
    {
        if (col < 1) throw new ArgumentOutOfRangeException(nameof(col));

        Span<char> buffer = stackalloc char[8];
        int at = buffer.Length;
        while (col > 0)
        {
            col--;
            buffer[--at] = (char)('A' + col % 26);
            col /= 26;
        }

        return new string(buffer[at..]);
    }

    /// <summary>Spell a one-based (column, row) pair as an A1 reference.</summary>
    public static string Name(int col, int row) => Column(col) + row.ToString();
}
