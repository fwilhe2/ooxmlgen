using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X4 — the cell's appearance. <c>cellXfs[s]</c> indexes a font, a fill, a border, an
/// alignment and a number format, and each of the five maps onto a piece of ODF's
/// <c>CellStyle</c> whose values are ODF's own strings.
/// </summary>
public static class CellStyles
{
    public static void Generate(Corpus corpus)
    {
        Fonts(corpus);
        Colours(corpus);
        Fills(corpus);
        Borders(corpus);
        Alignment(corpus);
        Named(corpus);
    }

    static void Fonts(Corpus corpus)
    {
        var spec = New("styles/fonts.xlsx", "X4", "weight, slant, underline, strike, size, script position");
        spec.Drops(Dropped.FontFamily, 6);
        spec.Note("The font *family* is deliberately not carried: grind's style model has no slot " +
                  "for one (§5.4), so every `<name val=\"…\"/>` in this file is a counted drop " +
                  "rather than a lost detail nobody noticed.");
        spec.Note("Superscript and subscript are `vertAlign` on the font in Excel and a text " +
                  "position in ODF — a different axis from the cell's vertical alignment, which " +
                  "styles/alignment.xlsx covers separately.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Fonts");
            tab.Text("A1", "font").Text("B1", "sample");

            var cases = new (string Label, Font Font, string? Note)[]
            {
                ("default", new Font(), null),
                ("bold", new Font(new Bold()), "font_weight: \"bold\""),
                ("italic", new Font(new Italic()), "font_style: \"italic\""),
                ("bold italic", new Font(new Bold(), new Italic()), null),
                ("underline single", new Font(new Underline()), "no `val` means single"),
                ("underline double", new Font(new Underline { Val = UnderlineValues.Double }), null),
                ("underline accounting", new Font(new Underline { Val = UnderlineValues.SingleAccounting }),
                 "an accounting underline spans the cell, not the text — no ODF spelling"),
                ("strikethrough", new Font(new Strike()), null),
                ("size 8", new Font(new FontSize { Val = 8 }), "font_size: \"8pt\""),
                ("size 11", new Font(new FontSize { Val = 11 }), "the default in every workbook Excel creates"),
                ("size 24", new Font(new FontSize { Val = 24 }), null),
                ("size 10.5", new Font(new FontSize { Val = 10.5 }), "half-point sizes are legal"),
                ("superscript", new Font(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript }), null),
                ("subscript", new Font(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript }), null),
                ("named Calibri", new Font(new FontName { Val = "Calibri" }, new FontFamilyNumbering { Val = 2 }),
                 "the family is dropped and counted"),
                ("named Courier New", new Font(new FontName { Val = "Courier New" }, new FontFamilyNumbering { Val = 3 }), null),
                ("named Comic Sans MS", new Font(new FontName { Val = "Comic Sans MS" }), null),
                ("named 日本語フォント", new Font(new FontName { Val = "ＭＳ ゴシック" }, new FontCharSet { Val = 128 }),
                 "a non-ASCII family name with a charset attribute"),
                ("scheme minor", new Font(new FontScheme { Val = FontSchemeValues.Minor }),
                 "the font is named by the theme rather than literally"),
                ("everything at once",
                 new Font(new Bold(), new Italic(), new Underline { Val = UnderlineValues.Double },
                          new Strike(), new FontSize { Val = 14 }, new FontName { Val = "Georgia" }), null),
            };

            int row = 2;
            foreach ((string label, Font font, string? note) in cases)
            {
                uint style = book.Xf(new CellFormat { FontId = book.FontId(font), ApplyFont = true });
                tab.Text($"A{row}", label);
                tab.Text($"B{row}", "Sample 123", style);
                if (note is not null) tab.Expect(note: note);
                row++;
            }
        }));
    }

    static void Colours(Corpus corpus)
    {
        var spec = New("styles/colors.xlsx", "X4", "rgb, indexed, theme and tint — the four ways to say a colour");
        spec.Drops(Dropped.ThemeColor, 14);
        spec.Note("Four spellings, and only the first is self-contained. `indexed` needs " +
                  "ECMA-376 §18.8.27's legacy palette; `theme` needs xl/theme/theme1.xml; `tint` " +
                  "needs the theme *and* ECMA's tint formula applied to it in HSL space.");
        spec.Note("The theme index is not the order the colours appear in the scheme: 0 is lt1 and " +
                  "1 is dk1, so the light and dark of each pair are swapped relative to the XML. " +
                  "A reader that maps by document order gets white and black the wrong way round " +
                  "and everything from accent1 on exactly right, which is the worst possible way " +
                  "to be wrong.");
        spec.Note("`rgb` is ARGB. The alpha byte is dropped — ODF's colour is six hex digits — and " +
                  "Excel writes FF there in essentially every file, so nothing real is lost.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            book.AddTheme();
            Tab tab = book.Sheet("Colours");
            tab.Text("A1", "spelling").Text("B1", "sample").Text("C1", "expected");

            int row = 2;

            void Row(string label, Color colour, string? expected, string? note = null)
            {
                uint style = book.Xf(new CellFormat
                {
                    FontId = book.FontId(new Font(colour, new Bold())),
                    ApplyFont = true,
                });
                tab.Text($"A{row}", label);
                tab.Text($"B{row}", "Sample", style);
                if (expected is not null) tab.Expect(note: $"color: \"{expected}\"");
                if (note is not null) tab.Expect(note: note);
                tab.Text($"C{row}", expected ?? "—");
                row++;
            }

            Row("rgb opaque red", new Color { Rgb = "FFFF0000" }, "#ff0000");
            Row("rgb opaque blue", new Color { Rgb = "FF0000FF" }, "#0000ff");
            Row("rgb half alpha", new Color { Rgb = "800000FF" }, "#0000ff",
                "the alpha byte is dropped, so this is the same colour as the row above");
            Row("rgb six digits", new Color { Rgb = "00FF00" }, "#00ff00",
                "some producers write six digits rather than eight — legal, and the alpha is implied");
            Row("auto", new Color { Auto = true }, null,
                "`auto` means the system's foreground; there is no fixed colour to carry");

            // The legacy indexed palette. §18.8.27 fixes the first 64; 64 and 65 are the system
            // foreground and background and have no entry at all.
            var indexed = new (uint Index, string? Hex, string Note)[]
            {
                (8, "#000000", "index 8 is black — the palette restarts at 8 after a duplicated first eight"),
                (9, "#ffffff", "white"),
                (10, "#ff0000", "red"),
                (12, "#0000ff", "blue"),
                (13, "#ffff00", "yellow"),
                (64, null, "the system foreground: no palette entry, resolve or drop"),
                (65, null, "the system background: likewise"),
            };

            foreach ((uint index, string? hex, string note) in indexed)
                Row($"indexed {index}", new Color { Indexed = index }, hex, note);

            for (uint theme = 0; theme < 12; theme++)
            {
                Row($"theme {theme}", new Color { Theme = theme }, "#" + Book.ThemeColours[theme].ToLowerInvariant(),
                    theme switch
                    {
                        0 => "lt1 — white, despite dk1 coming first in the scheme",
                        1 => "dk1 — black. Swap these two and every other index still looks right",
                        2 => "lt2",
                        3 => "dk2",
                        10 => "hlink",
                        11 => "folHlink",
                        _ => $"accent{theme - 3}",
                    });
            }

            var tints = new (double Tint, string Note)[]
            {
                (0.0, "no tint: the theme colour unchanged"),
                (0.5, "positive tint lightens, in HSL luminance, by the given fraction of the way to white"),
                (-0.5, "negative tint darkens, by the fraction of the way to black"),
                (0.9999, "very nearly white"),
                (-0.9999, "very nearly black"),
            };

            foreach ((double tint, string note) in tints)
                Row($"theme 4 tint {tint}", new Color { Theme = 4, Tint = tint }, null,
                    note + ". ECMA's formula applies to luminance, not to the RGB channels — " +
                    "tinting each channel by the same fraction is the common wrong answer");
        }));
    }

    static void Fills(Corpus corpus)
    {
        var spec = New("styles/fills.xlsx", "X4", "solid fills, all eighteen patterns, and a gradient");
        spec.Note("Only `solid` carries: its `fgColor` is the cell background and that is the whole " +
                  "translation. Every other pattern is two colours plus a texture, which ODF's cell " +
                  "background cannot say, and a gradient is not even two flat colours.");
        spec.Note("A pattern fill's *foreground* is the pattern's ink and its *background* is the " +
                  "paper — so for `solid`, the colour that matters is `fgColor` and not, as the " +
                  "names suggest, `bgColor`.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Fills");
            tab.Text("A1", "pattern").Text("B1", "sample");

            uint solid = book.Xf(new CellFormat
            {
                FillId = book.FillId(new Fill(new PatternFill(
                    new ForegroundColor { Rgb = "FFFFFF00" }) { PatternType = PatternValues.Solid })),
                ApplyFill = true,
            });
            tab.Text("A2", "solid yellow").Text("B2", "sample", solid)
               .Expect(note: "background: \"#ffff00\" — the only fill with a clean translation");

            int row = 3;
            var patterns = new[]
            {
                PatternValues.None, PatternValues.MediumGray, PatternValues.DarkGray,
                PatternValues.LightGray, PatternValues.DarkHorizontal, PatternValues.DarkVertical,
                PatternValues.DarkDown, PatternValues.DarkUp, PatternValues.DarkGrid,
                PatternValues.DarkTrellis, PatternValues.LightHorizontal, PatternValues.LightVertical,
                PatternValues.LightDown, PatternValues.LightUp, PatternValues.LightGrid,
                PatternValues.LightTrellis, PatternValues.Gray125, PatternValues.Gray0625,
            };

            foreach (var pattern in patterns)
            {
                uint style = book.Xf(new CellFormat
                {
                    FillId = book.FillId(new Fill(new PatternFill(
                        new ForegroundColor { Rgb = "FF3366CC" },
                        new BackgroundColor { Rgb = "FFFFFFFF" })
                    { PatternType = pattern })),
                    ApplyFill = true,
                });
                tab.Text($"A{row}", pattern.ToString());
                tab.Text($"B{row}", "sample", style)
                   .Expect(note: "a two-colour pattern: dropped and counted, not approximated by its foreground");
                row++;
            }

            uint gradient = book.Xf(new CellFormat
            {
                FillId = book.FillId(new Fill(new GradientFill(
                    new GradientStop(new Color { Rgb = "FFFF0000" }) { Position = 0 },
                    new GradientStop(new Color { Rgb = "FF0000FF" }) { Position = 1 })
                { Degree = 90 })),
                ApplyFill = true,
            });
            tab.Text($"A{row}", "gradient").Text($"B{row}", "sample", gradient)
               .Expect(note: "a linear gradient — not a colour at all");
        }));
    }

    static void Borders(Corpus corpus)
    {
        var spec = New("styles/borders.xlsx", "X4", "all thirteen border styles, every side, and diagonals");
        spec.Note("Excel names a border style; ODF spells one as a width, a style and a colour — " +
                  "`\"0.5pt solid #000000\"`. The name-to-width table is a measurement against the " +
                  "oracle rather than a number the spec states, which is why every style gets a " +
                  "cell here rather than a representative few.");
        spec.Note("`diagonalUp` and `diagonalDown` are attributes on the border, not sides of it: " +
                  "the `<diagonal>` element supplies the line and the two flags say which way it " +
                  "runs. Both may be set at once, giving a cross.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Borders");
            tab.Text("A1", "style").Text("B1", "all four sides").Text("C1", "left only");

            var styles = new[]
            {
                BorderStyleValues.Thin, BorderStyleValues.Medium, BorderStyleValues.Thick,
                BorderStyleValues.Double, BorderStyleValues.Hair, BorderStyleValues.Dotted,
                BorderStyleValues.Dashed, BorderStyleValues.DashDot, BorderStyleValues.DashDotDot,
                BorderStyleValues.MediumDashed, BorderStyleValues.MediumDashDot,
                BorderStyleValues.MediumDashDotDot, BorderStyleValues.SlantDashDot,
            };

            int row = 2;
            foreach (var style in styles)
            {
                var colour = new Color { Rgb = "FF000000" };
                uint all = book.Xf(new CellFormat
                {
                    BorderId = book.BorderId(new Border(
                        new LeftBorder(new Color { Rgb = "FF000000" }) { Style = style },
                        new RightBorder(new Color { Rgb = "FF000000" }) { Style = style },
                        new TopBorder(new Color { Rgb = "FF000000" }) { Style = style },
                        new BottomBorder(new Color { Rgb = "FF000000" }) { Style = style },
                        new DiagonalBorder())),
                    ApplyBorder = true,
                });
                uint left = book.Xf(new CellFormat
                {
                    BorderId = book.BorderId(new Border(
                        new LeftBorder(colour) { Style = style },
                        new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
                    ApplyBorder = true,
                });

                tab.Text($"A{row}", style.ToString());
                tab.Text($"B{row}", "box", all);
                tab.Text($"C{row}", "left", left)
                   .Expect(note: "one side only — the other three must not inherit it");
                row++;
            }

            uint none = book.Xf(new CellFormat
            {
                BorderId = book.BorderId(new Border(
                    new LeftBorder { Style = BorderStyleValues.None },
                    new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
                ApplyBorder = true,
            });
            tab.Text($"A{row}", "none").Text($"B{row}", "explicit none", none)
               .Expect(note: "an explicit `none` is not the same as no element: it overrides an inherited border");
            row++;

            uint coloured = book.Xf(new CellFormat
            {
                BorderId = book.BorderId(new Border(
                    new LeftBorder(new Color { Rgb = "FFFF0000" }) { Style = BorderStyleValues.Thick },
                    new RightBorder(new Color { Indexed = 12 }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color { Theme = 4 }) { Style = BorderStyleValues.Medium },
                    new BottomBorder(new Color { Theme = 4, Tint = -0.25 }) { Style = BorderStyleValues.Dashed },
                    new DiagonalBorder())),
                ApplyBorder = true,
            });
            tab.Text($"A{row}", "four colours, four spellings").Text($"B{row}", "mixed", coloured)
               .Expect(note: "rgb, indexed, theme and tinted-theme on the four sides of one cell");
            row++;

            foreach ((string label, bool up, bool down) in new[]
                     { ("diagonal up", true, false), ("diagonal down", false, true), ("cross", true, true) })
            {
                uint style = book.Xf(new CellFormat
                {
                    BorderId = book.BorderId(new Border(
                        new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(),
                        new DiagonalBorder(new Color { Rgb = "FF008000" }) { Style = BorderStyleValues.Thin })
                    {
                        DiagonalUp = up,
                        DiagonalDown = down,
                    }),
                    ApplyBorder = true,
                });
                tab.Text($"A{row}", label).Text($"B{row}", "diag", style)
                   .Expect(note: "a diagonal rule, which ODF has no cell-border slot for");
                row++;
            }
        }));
    }

    static void Alignment(Corpus corpus)
    {
        var spec = New("styles/alignment.xlsx", "X4", "horizontal, vertical, wrap, indent, rotation, shrink");
        spec.Note("`textRotation` is degrees 0–90 counter-clockwise, then 91–180 meaning 1–90 " +
                  "*clockwise*, and 255 meaning stacked vertically — three encodings in one " +
                  "attribute, and 255 is not a rotation at all.");
        spec.Note("`centerContinuous` centres across the cells to its right without merging them. " +
                  "It looks like a merge and is not one, which is worth a fixture on its own.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Alignment");
            tab.Text("A1", "alignment").Text("B1", "sample");

            int row = 2;

            void Row(string label, Alignment alignment, string? note = null)
            {
                uint style = book.Xf(new CellFormat(alignment) { ApplyAlignment = true });
                tab.Text($"A{row}", label);
                tab.Text($"B{row}", "The quick brown fox", style);
                if (note is not null) tab.Expect(note: note);
                row++;
            }

            foreach (var horizontal in new[]
                     {
                         HorizontalAlignmentValues.General, HorizontalAlignmentValues.Left,
                         HorizontalAlignmentValues.Center, HorizontalAlignmentValues.Right,
                         HorizontalAlignmentValues.Fill, HorizontalAlignmentValues.Justify,
                         HorizontalAlignmentValues.CenterContinuous, HorizontalAlignmentValues.Distributed,
                     })
            {
                Row($"horizontal {horizontal}", new Alignment { Horizontal = horizontal },
                    horizontal == HorizontalAlignmentValues.CenterContinuous
                        ? "centres across the following cells without merging them"
                        : horizontal == HorizontalAlignmentValues.Fill
                            ? "repeats the text to fill the width — not an alignment at all"
                            : null);
            }

            foreach (var vertical in new[]
                     {
                         VerticalAlignmentValues.Top, VerticalAlignmentValues.Center,
                         VerticalAlignmentValues.Bottom, VerticalAlignmentValues.Justify,
                         VerticalAlignmentValues.Distributed,
                     })
            {
                Row($"vertical {vertical}", new Alignment { Vertical = vertical });
            }

            Row("wrap", new Alignment { WrapText = true }, "wrap: true");
            Row("shrink to fit", new Alignment { ShrinkToFit = true }, "shrink and wrap are mutually exclusive");
            Row("indent 1", new Alignment { Horizontal = HorizontalAlignmentValues.Left, Indent = 1 },
                "indent is counted in units of the character width, not in points");
            Row("indent 5", new Alignment { Horizontal = HorizontalAlignmentValues.Left, Indent = 5 });
            Row("rotate 45", new Alignment { TextRotation = 45 }, "45 degrees counter-clockwise");
            Row("rotate 90", new Alignment { TextRotation = 90 }, "straight up");
            Row("rotate 135", new Alignment { TextRotation = 135 },
                "135 means 45 degrees *clockwise*: the second encoding hiding in the same attribute");
            Row("rotate 180", new Alignment { TextRotation = 180 }, "90 degrees clockwise");
            Row("stacked", new Alignment { TextRotation = 255 },
                "255 is not a rotation: it stacks the characters vertically, one under the next");
            Row("everything", new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center,
                WrapText = true,
                TextRotation = 30,
                Indent = 2,
            });
        }));
    }

    static void Named(Corpus corpus)
    {
        var spec = New("styles/named-styles.xlsx", "X4", "cellStyleXfs, and the applyX flags that switch inheritance off");
        spec.Note("A `cellXfs` entry points at a `cellStyleXfs` entry through `xfId`, and each " +
                  "`applyX` flag says whether the cell's own value for that group wins or the " +
                  "named style's does. A reader that ignores `xfId` gets the direct formatting " +
                  "right and every named style wrong.");
        spec.Note("Row 5 is the case that separates the two readings: the cell sets bold but " +
                  "`applyFont=\"0\"`, so the *style's* font is what applies — not the cell's.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint headingFont = book.FontId(new Font(new Bold(), new FontSize { Val = 16 },
                                                    new Color { Rgb = "FF1F4E79" }));
            uint noteFont = book.FontId(new Font(new Italic(), new Color { Rgb = "FF808080" }));
            uint plainFont = book.FontId(new Font(new Bold()));

            uint headingXf = book.NamedStyle("Heading 1",
                new CellFormat { FontId = headingFont, ApplyFont = true }, builtinId: 16);
            uint noteXf = book.NamedStyle("Note",
                new CellFormat { FontId = noteFont, ApplyFont = true }, builtinId: 10);

            Tab tab = book.Sheet("Named");

            uint heading = book.Xf(new CellFormat
            {
                FontId = headingFont,
                FormatId = headingXf,
                ApplyFont = true,
            });
            uint noteStyle = book.Xf(new CellFormat { FontId = noteFont, FormatId = noteXf, ApplyFont = true });
            uint overridden = book.Xf(new CellFormat
            {
                FontId = plainFont,
                FormatId = headingXf,
                ApplyFont = false,
            });

            tab.Text("A1", "Quarterly report", heading)
               .Expect(note: "a cell using the named style 'Heading 1' and repeating its font directly");
            tab.Text("A2", "a note about the report", noteStyle)
               .Expect(note: "the named style 'Note'");
            tab.Text("A3", "plain", 0).Expect(note: "the default style, xfId 0");
            tab.Text("A4", "direct bold", book.Xf(new CellFormat { FontId = plainFont, ApplyFont = true }))
               .Expect(note: "direct formatting with no named style behind it");
            tab.Text("A5", "applyFont off", overridden)
               .Expect(note: "the cell names a bold font and sets applyFont=\"0\", so the Heading 1 " +
                             "style's font wins. A reader that reads fontId and ignores the flag " +
                             "renders this bold rather than as a heading");
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "styles", Milestone = milestone, Covers = covers };
}
