using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X3 — number formats. Excel spells one as a code string; ODF spells it as an ordered
/// sequence of parts. This family is the input side of that translation, and the reason it
/// has to happen before a date is a date at all.
/// </summary>
public static class NumberFormats
{
    public static void Generate(Corpus corpus)
    {
        Builtins(corpus);
        Sections(corpus);
        Numeric(corpus);
        Currency(corpus);
        DateTimeCodes(corpus);
        Conditions(corpus);
    }

    /// <summary>
    /// ECMA-376 §18.8.30's built-in ids, each actually used by a cell.
    /// </summary>
    static void Builtins(Corpus corpus)
    {
        var spec = New("numfmt/builtins.xlsx", "X3", "every built-in format id that ECMA-376 assigns");
        spec.Note("A built-in id carries no code string in the file: the reader has to know the " +
                  "table. §18.8.30 lists id 14 as `mm-dd-yy`, and Excel renders it in the user's " +
                  "locale rather than as written — so mapping these by meaning rather than by " +
                  "literal code is what keeps a converted document looking like what its author saw.");
        spec.Note("`display` is deliberately unset on this fixture. The whole point of a built-in " +
                  "is that its rendering is locale-dependent, so asserting one spelling of it would " +
                  "assert a locale.");
        spec.Note("Ids 5–8, 23–36, 41–44 and 50–58 are reserved or locale-specific and are not " +
                  "written here; the gaps in column A are the spec's, not an omission.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Builtins");
            tab.Text("A1", "id").Text("B1", "code per §18.8.30").Text("C1", "value");

            // (id, the code the spec lists, the value to show it with, the kind it implies)
            var builtins = new (uint Id, string Code, string Value, ValueKind Kind)[]
            {
                (0, "General", "1234.5678", ValueKind.Number),
                (1, "0", "1234.5678", ValueKind.Number),
                (2, "0.00", "1234.5678", ValueKind.Number),
                (3, "#,##0", "1234.5678", ValueKind.Number),
                (4, "#,##0.00", "1234.5678", ValueKind.Number),
                (9, "0%", "0.1234", ValueKind.Percentage),
                (10, "0.00%", "0.1234", ValueKind.Percentage),
                (11, "0.00E+00", "1234.5678", ValueKind.Number),
                (12, "# ?/?", "1234.5678", ValueKind.Number),
                (13, "# ??/??", "1234.5678", ValueKind.Number),
                (14, "mm-dd-yy", "45368", ValueKind.Date),
                (15, "d-mmm-yy", "45368", ValueKind.Date),
                (16, "d-mmm", "45368", ValueKind.Date),
                (17, "mmm-yy", "45368", ValueKind.Date),
                (18, "h:mm AM/PM", "45368.75", ValueKind.Time),
                (19, "h:mm:ss AM/PM", "45368.75", ValueKind.Time),
                (20, "h:mm", "45368.75", ValueKind.Time),
                (21, "h:mm:ss", "45368.75", ValueKind.Time),
                (22, "m/d/yy h:mm", "45368.75", ValueKind.Date),
                (37, "#,##0 ;(#,##0)", "-1234.5678", ValueKind.Number),
                (38, "#,##0 ;[Red](#,##0)", "-1234.5678", ValueKind.Number),
                (39, "#,##0.00;(#,##0.00)", "-1234.5678", ValueKind.Number),
                (40, "#,##0.00;[Red](#,##0.00)", "-1234.5678", ValueKind.Number),
                (45, "mm:ss", "45368.75", ValueKind.Time),
                (46, "[h]:mm:ss", "45368.75", ValueKind.Time),
                (47, "mmss.0", "45368.75", ValueKind.Time),
                (48, "##0.0E+0", "1234.5678", ValueKind.Number),
                (49, "@", "1234.5678", ValueKind.Number),
            };

            int row = 2;
            foreach ((uint id, string code, string value, ValueKind kind) in builtins)
            {
                uint style = book.BuiltinFormat(id);
                tab.Number($"A{row}", id);
                tab.Text($"B{row}", code);
                tab.Number($"C{row}", value, style)
                   .Expect(kind: kind, note: $"built-in id {id}; the file carries the id alone, never the code");
                row++;
            }
        }));
    }

    /// <summary>The section rule: one code, up to four branches, chosen by the value's sign.</summary>
    static void Sections(Corpus corpus)
    {
        var spec = New("numfmt/sections.xlsx", "X3", "one to four sections, and the colours attached to them");
        spec.Note("A format code is up to four sections separated by `;`, applied to positive, " +
                  "negative, zero and text in that order. ODF spells the same thing as a base " +
                  "format plus `style:map` branches, which is what makes this a translation rather " +
                  "than a parse-and-drop.");
        spec.Note("The fourth section applies to *text* in the cell, which is the one most readers " +
                  "forget: a cell holding a string still has a number format, and it still selects.");
        spec.Note("Each format is applied to five values — positive, negative, zero, a string, and " +
                  "a value that exercises the rounding — so a comparison sees every branch.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Sections");
            tab.Text("A1", "code").Text("B1", "1234.5").Text("C1", "-1234.5")
               .Text("D1", "0").Text("E1", "text");

            var codes = new (string Code, string[] Display, string? Note)[]
            {
                ("0.00", ["1234.50", "-1234.50", "0.00", "text"],
                 "one section: negatives take the same format with a minus in front"),
                ("0.00;(0.00)", ["1234.50", "(1234.50)", "0.00", "text"],
                 "two sections: the negative branch supplies its own sign, so the minus goes"),
                ("0.00;[Red]-0.00", ["1234.50", "-1234.50", "0.00", "text"],
                 "a colour on the negative branch — the colour is the branch's, not the value's"),
                ("#,##0.00;[Red](#,##0.00);\"—\"", ["1,234.50", "(1,234.50)", "—", "text"],
                 "three sections: zero gets a literal em dash rather than a number"),
                ("#,##0.00;[Red](#,##0.00);\"—\";[Blue]@", ["1,234.50", "(1,234.50)", "—", "text"],
                 "all four: the last one applies to the string in column E"),
                ("[Green]0.0;[Magenta]0.0;[Cyan]0.0;[Yellow]@", ["1234.5", "1234.5", "0.0", "text"],
                 "a colour on every branch, including the text one"),
            };

            int row = 2;
            foreach ((string code, string[] display, string? note) in codes)
            {
                uint style = book.Format(code);
                tab.Text($"A{row}", code);
                tab.Number($"B{row}", "1234.5", style).Expect(display: display[0]);
                if (note is not null) tab.Expect(note: note);
                tab.Number($"C{row}", "-1234.5", style).Expect(display: display[1]);
                tab.Number($"D{row}", "0", style).Expect(display: display[2]);
                tab.Text($"E{row}", "text", style).Expect(display: display[3]);
                row++;
            }
        }));
    }

    /// <summary>Custom numeric codes: digits, grouping, literals, and the width tricks.</summary>
    static void Numeric(Corpus corpus)
    {
        var spec = New("numfmt/custom-numeric.xlsx", "X3", "custom codes: digits, grouping, padding, science, fractions");
        spec.Note("`_` consumes the character after it and leaves a space that wide — used to make " +
                  "positives line up with parenthesised negatives. `*` repeats the character after " +
                  "it to fill the column. Neither has an ODF spelling worth inventing; both have to " +
                  "be *parsed* regardless, because getting their argument wrong shifts everything " +
                  "after them.");
        spec.Note("Fractions (`# ?/?`) and scientific notation are both listed in grind's plan as " +
                  "possibly having no `numfmt::Part` to land on. If they are dropped, they are " +
                  "dropped knowingly — which is what this fixture is for.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Numeric");
            tab.Text("A1", "code").Text("B1", "value").Text("C1", "expected display");

            var cases = new (string Code, string Value, string Display, string? Note)[]
            {
                ("0", "3.7", "4", "no decimals: the value is rounded for display, not changed"),
                ("0.000", "3.7", "3.700", "trailing zeros are forced"),
                ("#.###", "3.7", "3.7", "`#` drops a digit that is not there; `0` keeps it"),
                ("#,##0", "1234567", "1,234,567", "grouping"),
                ("#,##0,", "1234567", "1,235", "a trailing comma divides by a thousand — one comma, two meanings"),
                ("#,##0,,\" M\"", "1234567890", "1,235 M", "two of them, plus a literal suffix"),
                ("0%", "0.1234", "12%", "percent multiplies by a hundred as well as appending a sign"),
                ("0.00%", "0.1234", "12.34%", null),
                ("0.00E+00", "12345.678", "1.23E+04", "scientific"),
                ("##0.0E+0", "12345.678", "12.3E+3", "engineering: the exponent moves in threes"),
                ("# ?/?", "3.75", "3 3/4", "a fraction with one denominator digit"),
                ("# ??/??", "3.14159", "3 14/99",
                 "two digits caps the denominator at 99, so the best available approximation of pi " +
                 "is 14/99 and not the 16/113 an unbounded search would find. The digit count is a " +
                 "constraint on the answer, not a hint about it"),
                ("# ?/8", "3.75", "3 6/8", "a fixed denominator, unreduced"),
                ("\\$0.00", "5", "$5.00", "an escaped literal"),
                ("\"total: \"0", "5", "total: 5", "a quoted literal"),
                ("0.00_);(0.00)", "5", "5.00 ", "`_)` leaves a space as wide as `)` so the columns line up"),
                ("0.00_);(0.00)", "-5", "(5.00)", "and the negative branch uses the real bracket"),
                ("*-0", "5", "5", "`*-` fills the *rest of the column* with dashes, so how many there " +
                                  "are is a fact about the column width rather than about the value. " +
                                  "`display` here is the text without the fill, which is what a " +
                                  "width-free rendering produces and the only part worth asserting"),
                ("@", "x", "x", "the text placeholder on its own"),
                ("\"<\"@\">\"", "x", "<x>", "literals around the text placeholder"),
                ("000000", "42", "000042", "zero padding to a fixed width"),
                ("0.0;;", "0", "", "an empty zero section hides zeros entirely"),
            };

            int row = 2;
            foreach ((string code, string value, string display, string? note) in cases)
            {
                uint style = book.Format(code);
                tab.Text($"A{row}", code);
                bool isText = !double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out _);
                if (isText) tab.Text($"B{row}", value, style);
                else tab.Number($"B{row}", value, style);
                tab.Expect(display: display);
                if (note is not null) tab.Expect(note: note);
                tab.Text($"C{row}", display);
                row++;
            }
        }));
    }

    /// <summary>Currency and locale tags, which is where a format code stops being just digits.</summary>
    static void Currency(Corpus corpus)
    {
        var spec = New("numfmt/currency-locale.xlsx", "X3", "[$symbol-locale] currency tags and locale-tagged dates");
        spec.Note("`[$€-407]` is a currency symbol plus an LCID. The symbol is the part that has " +
                  "to survive; the LCID says which locale's conventions Excel would render it with, " +
                  "and ODF carries a language and country on the format instead.");
        spec.Note("`[$-409]` with no symbol before the dash is a bare locale tag on a date format. " +
                  "A parser that assumes everything between `[$` and `-` is a symbol reads an empty " +
                  "one here rather than failing, which is the correct outcome and worth pinning.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Currency");
            tab.Text("A1", "code").Text("B1", "value");

            var cases = new (string Code, string Value, ValueKind Kind, string? Display, string? Note)[]
            {
                ("\"$\"#,##0.00", "1234.5", ValueKind.Currency, "$1,234.50", "the plain quoted symbol, no locale at all"),
                ("[$$-409]#,##0.00", "1234.5", ValueKind.Currency, null, "US dollar, LCID 409"),
                ("[$€-407]#,##0.00", "1234.5", ValueKind.Currency, null, "euro, German LCID — the symbol leads"),
                ("#,##0.00\\ [$€-407]", "1234.5", ValueKind.Currency, null, "and the same symbol trailing, which is how German actually writes it"),
                ("[$£-809]#,##0.00", "1234.5", ValueKind.Currency, null, "pound sterling"),
                ("[$¥-411]#,##0", "1234", ValueKind.Currency, null, "yen: no minor unit, so no decimals"),
                ("[$CHF-807]\\ #,##0.00", "1234.5", ValueKind.Currency, null, "a multi-character symbol"),
                ("[$-409]mmmm\\ d\\,\\ yyyy", "45368", ValueKind.Date, null, "a bare locale tag with no symbol before the dash"),
                ("[$-407]TT.MM.JJJJ", "45368", ValueKind.Date, null,
                 "German format letters. Excel stores localised codes in some files; the LCID is " +
                 "what says how to read them"),
            };

            int row = 2;
            foreach ((string code, string value, ValueKind kind, string? display, string? note) in cases)
            {
                uint style = book.Format(code);
                tab.Text($"A{row}", code);
                tab.Number($"B{row}", value, style).Expect(kind: kind);
                if (display is not null) tab.Expect(display: display);
                if (note is not null) tab.Expect(note: note);
                row++;
            }
        }));
    }

    /// <summary>Every date and time piece a code can be built from.</summary>
    static void DateTimeCodes(Corpus corpus)
    {
        var spec = New("numfmt/datetime-codes.xlsx", "X3", "y/m/d/h/s pieces, AM-PM, sub-seconds, elapsed");
        spec.Note("`m` is minutes when it follows `h` and months everywhere else. That one rule " +
                  "decides `h:m` and `m/d` both, and a parser that resolves it left-to-right gets " +
                  "`mm:ss` right by accident and `ss:mm` wrong for the same reason.");
        spec.Note("The value throughout is 45368.7541550926 — 2024-03-17 18:05:59.0 — chosen so " +
                  "every piece has a distinguishable value: the month is not the day, the hour is " +
                  "not the minute, and the sub-second digits are not zero.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("DateTime");
            tab.Text("A1", "code").Text("B1", "rendered").Text("C1", "note");

            const string value = "45368.7541550926";
            var cases = new (string Code, string Display, ValueKind Kind, string? Note)[]
            {
                ("yyyy", "2024", ValueKind.Date, null),
                ("yy", "24", ValueKind.Date, null),
                ("m", "3", ValueKind.Date, "bare `m` is the month, because nothing before it is an hour"),
                ("mm", "03", ValueKind.Date, null),
                ("mmm", "Mar", ValueKind.Date, null),
                ("mmmm", "March", ValueKind.Date, null),
                ("mmmmm", "M", ValueKind.Date, "one letter — the first of the month name"),
                ("d", "17", ValueKind.Date, null),
                ("dd", "17", ValueKind.Date, null),
                ("ddd", "Sun", ValueKind.Date, null),
                ("dddd", "Sunday", ValueKind.Date, null),
                ("yyyy-mm-dd", "2024-03-17", ValueKind.Date, "the ISO spelling, which Excel has no built-in for"),
                ("h", "18", ValueKind.Time, null),
                ("hh", "18", ValueKind.Time, null),
                ("h:mm", "18:05", ValueKind.Time, "`mm` after `h` is minutes"),
                ("mm:ss", "05:59", ValueKind.Time, "`mm` before `ss` is minutes as well"),
                ("h:mm:ss", "18:05:59", ValueKind.Time, null),
                ("h:mm AM/PM", "6:05 PM", ValueKind.Time, "the twelve-hour clock, with the marker"),
                ("h:mm A/P", "6:05 P", ValueKind.Time,
                 "the one-letter marker. Excel writes it in the case the code uses; LibreOffice " +
                 "renders `6:05 p` lowercase — measured 2026-09-14. The case is the implementation's, " +
                 "so a comparison should fold it rather than treat this cell as a mismatch"),
                ("ss.0", "59.0", ValueKind.Time, "tenths"),
                ("ss.000", "59.000", ValueKind.Time, "thousandths"),
                ("[h]:mm", "1088850:05", ValueKind.Time, "elapsed hours: the whole serial, not a clock reading"),
                ("[m]", "65331005", ValueKind.Time, "elapsed minutes"),
                ("[s]", "3919860359", ValueKind.Time, "elapsed seconds"),
                ("yyyy-mm-dd hh:mm:ss", "2024-03-17 18:05:59", ValueKind.Date, "date and time together"),
            };

            int row = 2;
            foreach ((string code, string display, ValueKind kind, string? note) in cases)
            {
                uint style = book.Format(code);
                tab.Text($"A{row}", code);
                tab.Number($"B{row}", value, style).Expect(kind: kind, display: display);
                if (note is not null) tab.Expect(note: note);
                if (note is not null) tab.Text($"C{row}", note);
                row++;
            }
        }));
    }

    /// <summary>Conditions, which are the part of the section rule that has no ODF spelling.</summary>
    static void Conditions(Corpus corpus)
    {
        var spec = New("numfmt/conditions.xlsx", "X3", "[>=100]-style conditions beyond the plain section rule");
        spec.Note("A section may carry a condition instead of taking its place in the " +
                  "positive/negative/zero order. ODF's `style:map` models a two-branch condition " +
                  "and the renderer already follows one level, so the first is close to free and " +
                  "the rest are dropped and counted rather than approximated.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Conditions");
            tab.Text("A1", "code").Text("B1", "50").Text("C1", "500").Text("D1", "-50");

            var codes = new (string Code, string? Note)[]
            {
                ("[>=100]#,##0;0.00", "one condition and a default: the shape `style:map` already models"),
                ("[>=100]\"big\";[<=-100]\"small\";0.00", "two conditions and a default — one level past what the renderer follows"),
                ("[<50][Red]0;[>500][Blue]0;[Green]0", "conditions and colours together"),
                ("[=0]\"zero\";General", "an equality condition"),
            };

            int row = 2;
            foreach ((string code, string? note) in codes)
            {
                uint style = book.Format(code);
                tab.Text($"A{row}", code);
                tab.Number($"B{row}", "50", style);
                if (note is not null) tab.Expect(note: note);
                tab.Number($"C{row}", "500", style);
                tab.Number($"D{row}", "-50", style);
                row++;
            }
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "numfmt", Milestone = milestone, Covers = covers };
}
