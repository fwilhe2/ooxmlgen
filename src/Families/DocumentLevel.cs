using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OoxmlGen.Sdk;

namespace OoxmlGen.Families;

/// <summary>
/// X5 — everything above the cell: names, sheet order and visibility, merges, filters, and
/// the parts whose whole contribution to a conversion is an honest line in the report.
/// </summary>
public static class DocumentLevel
{
    public static void Generate(Corpus corpus)
    {
        DefinedNames(corpus);
        Sheets(corpus);
        MergedCells(corpus);
        AutoFilters(corpus);
        Tables(corpus);
        Comments(corpus);
        ConditionalFormats(corpus);
        DataValidations(corpus);
        Protection(corpus);
        MacroEnabled(corpus);
        ExternalLink(corpus);
        Chart(corpus);
        PivotTable(corpus);
    }

    static void DefinedNames(Corpus corpus)
    {
        var spec = New("document/defined-names.xlsx", "X5", "global names, sheet-local names, and the _xlnm. built-ins");
        spec.Drops(Dropped.SheetLocalName, 2);
        spec.Note("A name with a `localSheetId` is scoped to one sheet, so two sheets may each " +
                  "define `Rate` and mean different things. The model's names are document-wide, " +
                  "which is why these are dropped rather than flattened: flattening them would " +
                  "silently pick one of the two.");
        spec.Note("`_xlnm.Print_Area` and `_xlnm.Print_Titles` are print settings wearing a name's " +
                  "clothes. They are not user names and should not arrive as if they were.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab data = book.Sheet("Data");
            Tab other = book.Sheet("Other");

            for (int row = 1; row <= 5; row++)
            {
                data.Number($"A{row}", row).Number($"B{row}", row * 1.5);
                other.Number($"A{row}", row * 100);
            }

            book.DefinedNames.AddRange(
            [
                new DefinedName { Name = "SingleCell", Text = "Data!$A$1" },
                new DefinedName { Name = "Range", Text = "Data!$A$1:$B$5" },
                new DefinedName { Name = "Constant", Text = "42" },
                new DefinedName { Name = "Expression", Text = "Data!$A$1*2+1" },
                new DefinedName { Name = "Multi_Area", Text = "Data!$A$1:$A$2,Data!$B$4:$B$5" },
                new DefinedName { Name = "CrossSheet", Text = "Other!$A$1" },
                new DefinedName { Name = "Name_With_Underscores", Text = "Data!$B$1" },
                new DefinedName { Name = "Unicode_Ω", Text = "Data!$B$2" },
                // Scoped to one sheet, which the model has no home for.
                new DefinedName { Name = "Rate", Text = "Data!$A$3", LocalSheetId = 0 },
                new DefinedName { Name = "Rate", Text = "Other!$A$3", LocalSheetId = 1 },
                // Print settings, not names.
                new DefinedName { Name = "_xlnm.Print_Area", Text = "Data!$A$1:$B$5", LocalSheetId = 0 },
                new DefinedName { Name = "_xlnm.Print_Titles", Text = "Data!$1:$1", LocalSheetId = 0 },
            ]);

            data.Formula("D1", "SUM(Range)", "22.5", "SUM(Range)")
                .Expect(note: "a formula naming a defined name: the name has to survive for this to mean anything");
            data.Formula("D2", "Constant*2", "84", "Constant*2");
            data.Formula("D3", "Rate", "3", null)
                .Expect(note: "the sheet-local name. Dropped, so this formula loses its referent " +
                              "and the report is the only thing that says so");
        }));
    }

    static void Sheets(Corpus corpus)
    {
        var spec = New("document/sheets.xlsx", "X5", "sheet order, hidden and veryHidden, and names ODF cannot spell");
        spec.Drops(Dropped.HiddenSheet, 2);
        spec.Note("A hidden sheet's *data* imports; it is the visibility that has nowhere to go. " +
                  "Dropping the data instead would be the one loss a user cannot detect.");
        spec.Note("`veryHidden` cannot be unhidden from Excel's UI at all — only from VBA. It is " +
                  "still just a sheet, and still full of cells.");
        spec.Note("The last two sheets carry names Excel permits and ODF does not. ODF forbids " +
                  "`[ ] * ? : / \\` and a leading or trailing apostrophe in a table name; Excel " +
                  "forbids a shorter list. A rename has to be deterministic — the same input giving " +
                  "the same output every run — and reported, and `expectName` is this fixture's " +
                  "suggestion rather than a requirement.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            var sheets = new (string Name, SheetStateValues? State, string? Expect, string Note)[]
            {
                ("First", null, null, "ordinary, and the first in the workbook"),
                ("Second", null, null, "ordinary"),
                ("Hidden", SheetStateValues.Hidden, null, "hidden: the data still imports"),
                ("VeryHidden", SheetStateValues.VeryHidden, null, "veryHidden: unreachable from the UI, still data"),
                ("With Spaces", null, null, "a space is legal in both"),
                ("Ünïcødé Ω 日本", null, null, "non-ASCII is legal in both"),
                ("O'Brien", null, null, "an apostrophe inside the name is legal in both"),
                ("A-Very-Long-Sheet-Name-Of-31c", null, null, "31 characters, which is Excel's maximum"),
                ("Has[Brackets]", null, "Has_Brackets_", "brackets: legal in Excel, forbidden in ODF"),
                ("Has/Slash", null, "Has_Slash", "a forward slash: likewise"),
            };

            uint index = 1;
            foreach ((string name, SheetStateValues? state, string? expect, string note) in sheets)
            {
                Tab tab = book.Sheet(name, state, expect);
                tab.Text("A1", note);
                tab.Number("B1", index);
                tab.Text("A2", $"this is sheet {index} of {sheets.Length}");
                index++;
            }
        }));
    }

    static void MergedCells(Corpus corpus)
    {
        var spec = New("document/merged-cells.xlsx", "X5", "merge ranges — counted, never moved");
        spec.Drops(Dropped.MergedCells, 7);
        spec.Note("The model carries no spans, so a merge is dropped. What must *not* happen is " +
                  "the values moving: in a merged range Excel keeps the value in the top-left cell " +
                  "and leaves the rest empty, and that is already exactly what an unmerged sheet " +
                  "looks like. Dropping the merge costs the appearance and nothing else.");
        spec.Note("A5:A7 is the case to check: its top-left is empty, so a reader that 'fills the " +
                  "range from the anchor' invents content rather than losing it.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint centred = book.Xf(new CellFormat(new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center,
            })
            { ApplyAlignment = true });

            Tab tab = book.Sheet("Merged");

            tab.Text("A1", "a horizontal merge across three columns", centred);
            tab.Text("A3", "vertical", centred);
            tab.Text("C3", "a block", centred);
            tab.Number("F1", 42, centred);
            tab.Text("A9", "outside every merge");

            tab.Merges = new MergeCells(
                new MergeCell { Reference = "A1:C1" },
                new MergeCell { Reference = "A3:A4" },
                new MergeCell { Reference = "C3:E5" },
                new MergeCell { Reference = "F1:F1" },
                new MergeCell { Reference = "A5:A7" },
                new MergeCell { Reference = "H1:I2" },
                new MergeCell { Reference = "A11:XFD11" })
            { Count = 7 };

            tab.Expect(note: "F1:F1 is a one-cell merge — legal, pointless, and written by real producers");
            tab.Text("A12", "H1:I2 is a merge over cells that do not exist at all")
               .Expect(note: "a merge does not imply a cell; the range may be entirely empty");
            tab.Text("A13", "A11:XFD11 merges an entire row")
               .Expect(note: "a whole-row merge is a range of 16384 columns and still not 16384 cells");
        }));
    }

    static void AutoFilters(Corpus corpus)
    {
        var spec = New("document/autofilter.xlsx", "X5", "autoFilter ranges, and filterColumn criteria");
        spec.Note("The *range* is carried: the model gained `filter.rs` after grind's import plan " +
                  "was first written, so `<autoFilter ref=\"A1:C9\"/>` maps onto `Filter::new` and " +
                  "is not a loss at all.");
        spec.Note("The *criteria* are a second question. A `<filterColumn>` holds discrete values, " +
                  "a custom comparison, a top-10 rule, a dynamic date band or a colour — and those " +
                  "carry where the model's own filter vocabulary has room and are counted where it " +
                  "does not.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab plain = book.Sheet("RangeOnly");
            Header(plain);
            plain.AutoFilter = new AutoFilter { Reference = "A1:C9" };
            plain.Text("E1", "a filter range with no criteria at all — the common case")
                 .Expect(note: "carried: Filter::new over A1:C9");

            Tab criteria = book.Sheet("WithCriteria");
            Header(criteria);
            criteria.AutoFilter = new AutoFilter(
                new FilterColumn(new Filters(
                    new Filter { Val = "north" },
                    new Filter { Val = "south" }))
                { ColumnId = 0 },
                new FilterColumn(new CustomFilters(
                    new CustomFilter { Operator = FilterOperatorValues.GreaterThan, Val = "100" },
                    new CustomFilter { Operator = FilterOperatorValues.LessThan, Val = "500" })
                { And = true })
                { ColumnId = 1 })
            { Reference = "A1:C9" };
            criteria.Text("E1", "discrete values on one column, a custom range on another");

            Tab top = book.Sheet("TopTen");
            Header(top);
            top.AutoFilter = new AutoFilter(
                new FilterColumn(new Top10 { Val = 3, Percent = false, Top = true }) { ColumnId = 1 })
            { Reference = "A1:C9" };
            top.Text("E1", "a top-10 rule — a criterion that depends on the data rather than describing it")
               .Expect(note: "nothing in the model's filter vocabulary expresses 'the largest three'");

            static void Header(Tab tab)
            {
                tab.Text("A1", "region").Text("B1", "amount").Text("C1", "date");
                string[] regions = ["north", "south", "east", "west", "north", "south", "east", "west"];
                for (int i = 0; i < regions.Length; i++)
                {
                    int row = i + 2;
                    tab.Text($"A{row}", regions[i]);
                    tab.Number($"B{row}", (i + 1) * 75);
                    tab.Number($"C{row}", 45300 + i * 7);
                }
            }
        }));
    }

    static void Tables(Corpus corpus)
    {
        var spec = New("document/tables.xlsx", "X5", "a real table part, and the structured references naming it");
        spec.Drops(Dropped.StructuredReference, 4);
        spec.Note("A structured reference needs a table to refer to, so this fixture builds one " +
                  "rather than writing `Table1[Amount]` into a workbook with no Table1 in it — a " +
                  "file Excel would refuse to open and an importer would be wrong to trust.");
        spec.Note("`[#This Row]`, `[#Headers]` and `[#All]` are the special item specifiers. Each " +
                  "means a different range and none of them is a cell reference, which is why the " +
                  "whole class is excluded rather than rewritten.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Sales");

            tab.Text("A1", "Region").Text("B1", "Amount").Text("C1", "Tax");
            string[] regions = ["north", "south", "east", "west"];
            for (int i = 0; i < regions.Length; i++)
            {
                int row = i + 2;
                tab.Text($"A{row}", regions[i]);
                tab.Number($"B{row}", (i + 1) * 100);
                tab.Formula($"C{row}", "Sales[[#This Row],[Amount]]*0.2", ((i + 1) * 20).ToString(), null)
                   .Expect(note: "[#This Row] — a structured reference to the same row of another column");
            }

            tab.Formula("E1", "SUM(Sales[Amount])", "1000", null)
               .Expect(note: "a whole-column structured reference");
            tab.Formula("E2", "COUNTA(Sales[#Headers])", "3", null)
               .Expect(note: "[#Headers] — the header row, which is not part of the data");
            tab.Formula("E3", "ROWS(Sales[#All])", "5", null)
               .Expect(note: "[#All] — headers and data together");
            tab.Formula("E4", "SUM(B2:B5)", "1000", "SUM([.B2:.B5])")
               .Expect(note: "the same sum written as an ordinary range: this one carries, and is " +
                             "the control for the four above");

            TableDefinitionPart table = tab.PartOf.AddNewPart<TableDefinitionPart>();
            Book.WriteXml(table, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       id="1" name="Sales" displayName="Sales" ref="A1:C5" totalsRowShown="0">
                  <autoFilter ref="A1:C5"/>
                  <tableColumns count="3">
                    <tableColumn id="1" name="Region"/>
                    <tableColumn id="2" name="Amount"/>
                    <tableColumn id="3" name="Tax"/>
                  </tableColumns>
                  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
                                  showRowStripes="1" showColumnStripes="0"/>
                </table>
                """);

            tab.TableParts = new TableParts(
                new TablePart { Id = tab.PartOf.GetIdOfPart(table) })
            { Count = 1 };
        }));
    }

    static void Comments(Corpus corpus)
    {
        var spec = New("document/comments.xlsx", "X5", "cell comments, their author list, and the VML that anchors them");
        spec.Drops(Dropped.Comment, 3);
        spec.Note("A comment lives in a separate part with its own author table, and is anchored " +
                  "to the sheet by a VML drawing from the 2007 era. Three parts and two " +
                  "vocabularies for a sticky note; all of it counted and none of it carried.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Commented");
            tab.Number("A1", 1).Number("B2", 2).Text("C3", "annotated");
            tab.Text("A5", "three cells carry comments: A1, B2 and C3");

            WorksheetCommentsPart comments = tab.PartOf.AddNewPart<WorksheetCommentsPart>();
            Book.WriteXml(comments, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <authors>
                    <author>Reviewer One</author>
                    <author>Reviewer Two</author>
                  </authors>
                  <commentList>
                    <comment ref="A1" authorId="0">
                      <text><r><t>a plain comment</t></r></text>
                    </comment>
                    <comment ref="B2" authorId="1">
                      <text>
                        <r><rPr><b/></rPr><t>bold</t></r>
                        <r><t xml:space="preserve"> and plain in one comment</t></r>
                      </text>
                    </comment>
                    <comment ref="C3" authorId="0">
                      <text><r><t>a comment on a text cell</t></r></text>
                    </comment>
                  </commentList>
                </comments>
                """);

            VmlDrawingPart vml = tab.PartOf.AddNewPart<VmlDrawingPart>();
            Book.WriteXml(vml, """
                <xml xmlns:v="urn:schemas-microsoft-com:vml"
                     xmlns:o="urn:schemas-microsoft-com:office:office"
                     xmlns:x="urn:schemas-microsoft-com:office:excel">
                  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
                  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
                    <v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>
                  </v:shapetype>
                  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:hidden"
                           fillcolor="#ffffe1" o:insetmode="auto">
                    <v:shadow on="t" color="black" obscured="t"/>
                    <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData>
                  </v:shape>
                </xml>
                """);

            tab.LegacyDrawing = new LegacyDrawing { Id = tab.PartOf.GetIdOfPart(vml) };
        }));
    }

    static void ConditionalFormats(Corpus corpus)
    {
        var spec = New("document/conditional-format.xlsx", "X5", "cellIs, expression, colour scale, data bar, icon set");
        spec.Drops(Dropped.ConditionalFormat, 6);
        spec.Note("Conditional formats are style *rules* rather than styles: what a cell looks " +
                  "like depends on its value at render time. The model has no rule engine, so all " +
                  "five kinds are counted. The cells keep their values and their direct formatting.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Conditional");
            for (int row = 1; row <= 10; row++)
            {
                tab.Number($"A{row}", row * 10);
                tab.Number($"B{row}", 55 - row * 5);
                tab.Number($"C{row}", row % 3);
            }

            tab.Text("E1", "A: cellIs, B: colour scale, C: expression")
               .Expect(note: "the values are ordinary; only the rules attached to them are not");

            // The dxf a rule points at lives in the stylesheet's `dxfs`, which this corpus does
            // not build — so the rules below reference dxfId 0 and the absence of a `dxfs` block
            // is itself a tolerance case.
            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("50"))
                {
                    Type = ConditionalFormatValues.CellIs,
                    Operator = ConditionalFormattingOperatorValues.GreaterThan,
                    Priority = 1,
                    FormatId = 0,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A10" } });

            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new ColorScale(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                    new Color { Rgb = "FFF8696B" },
                    new Color { Rgb = "FF63BE7B" }))
                {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = 2,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1:B10" } });

            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("MOD(C1,2)=0"))
                {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 3,
                    FormatId = 0,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "C1:C10" } });

            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new DataBar(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                    new Color { Rgb = "FF638EC6" }))
                {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = 4,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A10" } });

            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new IconSet(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "0" },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "33" },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "67" })
                { IconSetValue = IconSetValues.ThreeTrafficLights1 })
                {
                    Type = ConditionalFormatValues.IconSet,
                    Priority = 5,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1:B10" } });

            tab.ConditionalFormats.Add(new ConditionalFormatting(
                new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DuplicateValues,
                    Priority = 6,
                    FormatId = 0,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "C1:C10" } });
        }));
    }

    static void DataValidations(Corpus corpus)
    {
        var spec = New("document/data-validation.xlsx", "X5", "list, whole, decimal, date and custom validation");
        spec.Drops(Dropped.DataValidation, 5);
        spec.Note("Validation constrains what may be typed *next*, which is a property of an " +
                  "editing session rather than of a document's contents. Counted, and the cells " +
                  "keep whatever they already hold — including values that would fail the rule, " +
                  "because Excel lets those in through paste.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Validated");
            tab.Text("A1", "list").Text("B1", "whole").Text("C1", "decimal")
               .Text("D1", "date").Text("E1", "custom");

            tab.Text("A2", "red").Number("B2", 5).Number("C2", 0.5).Number("D2", 45368).Text("E2", "ok");
            tab.Text("A3", "purple").Number("B3", 999)
               .Expect(note: "a value the rule would reject, which paste can put there anyway");

            tab.Validations = new DataValidations(
                new DataValidation(new Formula1("\"red,green,blue\""))
                {
                    Type = DataValidationValues.List,
                    AllowBlank = true,
                    ShowInputMessage = true,
                    ShowErrorMessage = true,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A20" },
                },
                new DataValidation(new Formula1("1"), new Formula2("100"))
                {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    ErrorTitle = "Out of range",
                    Error = "Enter a whole number from 1 to 100.",
                    ShowErrorMessage = true,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:B20" },
                },
                new DataValidation(new Formula1("0"), new Formula2("1"))
                {
                    Type = DataValidationValues.Decimal,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "C2:C20" },
                },
                new DataValidation(new Formula1("45292"), new Formula2("45657"))
                {
                    Type = DataValidationValues.Date,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "D2:D20" },
                },
                new DataValidation(new Formula1("LEN(E2)<10"))
                {
                    Type = DataValidationValues.Custom,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "E2:E20" },
                })
            { Count = 5 };
        }));
    }

    static void Protection(Corpus corpus)
    {
        var spec = New("document/protection.xlsx", "X5", "sheet and workbook protection, and the locked flag");
        spec.Drops(Dropped.Protection, 2);
        spec.Note("Sheet protection is a UI lock with a hash beside it, not encryption: the file " +
                  "is a perfectly ordinary zip and every value in it is readable without the " +
                  "password. Dropping it loses nothing a reader was entitled to keep. An " +
                  "*encrypted* workbook is a different file entirely — see hostile/encrypted.xlsx.");
        spec.Note("`locked` on a cell's xf only matters while the sheet is protected, so the flag " +
                  "and the protection are one fact stored in two places.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            uint unlocked = book.Xf(new CellFormat(new Protection { Locked = false })
            {
                ApplyProtection = true,
            });
            uint hidden = book.Xf(new CellFormat(new Protection { Locked = true, Hidden = true })
            {
                ApplyProtection = true,
            });

            Tab tab = book.Sheet("Protected");
            tab.Text("A1", "locked by default", 0);
            tab.Text("A2", "explicitly unlocked", unlocked)
               .Expect(note: "the one cell an editor may change while the sheet is protected");
            tab.Formula("A3", "1+1", "2", "1+1", style: hidden)
               .Expect(value: "2", note: "`hidden` hides the *formula* from the formula bar; the value still shows");

            tab.Protection = new SheetProtection
            {
                Sheet = true,
                Objects = true,
                Scenarios = true,
                AlgorithmName = "SHA-512",
                HashValue = "0YLBGGIz+X4qpBpnUNIgRJhTLxbnkOpBrK2tVs0J9Dc=",
                SaltValue = "MhS3sMEZEPnlyRJ8CvqPvg==",
                SpinCount = 100000,
            };

            book.Protection = new WorkbookProtection
            {
                LockStructure = true,
                LockWindows = false,
                WorkbookAlgorithmName = "SHA-512",
                WorkbookHashValue = "1hLCHHJz+Y5rqCqoVOJhSKiUMycokPqCsL3uWt1K0Ed=",
                WorkbookSaltValue = "NiT4tNFAFQomzSK9DwrQwh==",
                WorkbookSpinCount = 100000,
            };
        }));
    }

    static void MacroEnabled(Corpus corpus)
    {
        var spec = New("document/macro-enabled.xlsm", "X5", ".xlsm — the same XML, plus a macro nobody runs");
        spec.Drops(Dropped.Macro);
        spec.Note("A macro-enabled workbook is not a different format: the cells, styles and " +
                  "formulas are byte-for-byte the same vocabulary. Only the content type of the " +
                  "package and one extra binary part differ.");
        spec.Note("`vbaProject.bin` is a CFB container holding compiled VBA. It is data to be " +
                  "counted and never, under any circumstances, executed — so the bytes here are " +
                  "a CFB header and filler, not a real project.");

        corpus.Emit(spec, path => Book.Create(path, spec, SpreadsheetDocumentType.MacroEnabledWorkbook, book =>
        {
            Tab tab = book.Sheet("Data");
            tab.Text("A1", "product").Text("B1", "price");
            tab.Text("A2", "widget").Number("B2", 9.99);
            tab.Text("A3", "gadget").Number("B3", 24.5);
            tab.Formula("B4", "SUM(B2:B3)", "34.49", "SUM([.B2:.B3])");
            tab.Text("A5", "a workbook with a macro in it, importing exactly like one without")
               .Expect(note: "the point of the fixture: the data must not be treated with suspicion");

            VbaProjectPart vba = book.Part.AddNewPart<VbaProjectPart>();
            byte[] cfb = new byte[1536];
            // The CFB signature, so the part looks like what it claims to be rather than zeros.
            ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            signature.CopyTo(cfb);
            using Stream stream = vba.GetStream(FileMode.Create);
            stream.Write(cfb);
        }));
    }

    static void ExternalLink(Corpus corpus)
    {
        var spec = New("document/external-link.xlsx", "X5", "a link to another workbook — recorded, never fetched");
        spec.Drops(Dropped.ExternalLink, 3);
        spec.Note("`[1]Sheet1!A1` refers into the first entry of `externalReferences`, whose part " +
                  "names a file on someone else's disk and caches the values last seen there. The " +
                  "cache is why the formulas below still have numbers.");
        spec.Note("A converter must never resolve one of these. Following a link means reading a " +
                  "path chosen by whoever sent the file, which is a different threat model from " +
                  "parsing bytes they sent — and there is no case where a batch conversion should " +
                  "touch the filesystem or the network on the document's instructions.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Linked");
            tab.Text("A1", "these formulas name a workbook that is not here");
            tab.Formula("B1", "[1]Sheet1!$A$1", "100", null)
               .Expect(note: "a single external cell; the value is Excel's cache of it");
            tab.Formula("B2", "SUM([1]Sheet1!$A$1:$A$3)", "600", null);
            tab.Formula("B3", "[1]!ExternalName", "42", null)
               .Expect(note: "a defined name in the other workbook, with no sheet at all");
            tab.Formula("B4", "SUM(B1:B2)", "700", "SUM([.B1:.B2])")
               .Expect(note: "a local formula over the cached values — this one carries");

            ExternalWorkbookPart external = book.Part.AddNewPart<ExternalWorkbookPart>();
            Book.WriteXml(external, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <externalBook r:id="rId1">
                    <sheetNames><sheetName val="Sheet1"/></sheetNames>
                    <definedNames><definedName name="ExternalName" refersTo="Sheet1!$C$1"/></definedNames>
                    <sheetDataSet>
                      <sheetData sheetId="0">
                        <row r="1"><cell r="A1"><v>100</v></cell></row>
                        <row r="2"><cell r="A2"><v>200</v></cell></row>
                        <row r="3"><cell r="A3"><v>300</v></cell></row>
                      </sheetData>
                    </sheetDataSet>
                  </externalBook>
                </externalLink>
                """);

            // The target is a path on a machine that is not this one. It exists to be recorded
            // and refused, so it is deliberately absolute and deliberately elsewhere.
            external.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                new Uri("file:///home/someone-else/Documents/Budget.xlsx"),
                "rId1");

            book.ExternalReferences.Add(new ExternalReference { Id = book.Part.GetIdOfPart(external) });
        }));
    }

    static void Chart(Corpus corpus)
    {
        var spec = New("document/chart.xlsx", "X5", "a drawing and a chart — DrawingML, counted rather than read");
        spec.Drops(Dropped.Chart).Drops(Dropped.Drawing);
        spec.Note("The model does have somewhere to put a chart — grind's `sheet/src/chart.rs` " +
                  "exists. What is missing is a reader for `xl/charts/chart1.xml`, which is " +
                  "DrawingML: a second vocabulary about the size of the whole import filter. So " +
                  "this is dropped for cost, not for want of a home, and the report should say so.");
        spec.Note("A chart arrives through two parts, not one: the sheet points at a *drawing*, " +
                  "and the drawing's graphic frame points at the chart. Counting one and missing " +
                  "the other is the easy mistake.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab tab = book.Sheet("Data");
            tab.Text("A1", "quarter").Text("B1", "revenue");
            string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
            for (int i = 0; i < quarters.Length; i++)
            {
                tab.Text($"A{i + 2}", quarters[i]);
                tab.Number($"B{i + 2}", 100 + i * 45);
            }

            tab.Text("D1", "the numbers are ordinary; the chart over them is the part with no home");

            DrawingsPart drawing = tab.PartOf.AddNewPart<DrawingsPart>();
            ChartPart chart = drawing.AddNewPart<ChartPart>();

            Book.WriteXml(chart, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart>
                    <c:title><c:overlay val="0"/></c:title>
                    <c:plotArea>
                      <c:layout/>
                      <c:barChart>
                        <c:barDir val="col"/>
                        <c:grouping val="clustered"/>
                        <c:varyColors val="0"/>
                        <c:ser>
                          <c:idx val="0"/>
                          <c:order val="0"/>
                          <c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
                          <c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>
                          <c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
                        </c:ser>
                        <c:axId val="111111111"/>
                        <c:axId val="222222222"/>
                      </c:barChart>
                      <c:catAx>
                        <c:axId val="111111111"/>
                        <c:scaling><c:orientation val="minMax"/></c:scaling>
                        <c:delete val="0"/>
                        <c:axPos val="b"/>
                        <c:crossAx val="222222222"/>
                      </c:catAx>
                      <c:valAx>
                        <c:axId val="222222222"/>
                        <c:scaling><c:orientation val="minMax"/></c:scaling>
                        <c:delete val="0"/>
                        <c:axPos val="l"/>
                        <c:crossAx val="111111111"/>
                      </c:valAx>
                    </c:plotArea>
                    <c:plotVisOnly val="1"/>
                  </c:chart>
                </c:chartSpace>
                """);

            Book.WriteXml(drawing, $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <xdr:twoCellAnchor>
                    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>21</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                    <xdr:graphicFrame macro="">
                      <xdr:nvGraphicFramePr>
                        <xdr:cNvPr id="2" name="Chart 1"/>
                        <xdr:cNvGraphicFramePr/>
                      </xdr:nvGraphicFramePr>
                      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                                   r:id="{drawing.GetIdOfPart(chart)}"/>
                        </a:graphicData>
                      </a:graphic>
                    </xdr:graphicFrame>
                    <xdr:clientData/>
                  </xdr:twoCellAnchor>
                </xdr:wsDr>
                """);

            tab.Drawing = new Drawing { Id = tab.PartOf.GetIdOfPart(drawing) };
        }));
    }

    static void PivotTable(Corpus corpus)
    {
        var spec = New("document/pivot-table.xlsx", "X5", "a pivot table, its cache, and the sheet it renders onto");
        spec.Drops(Dropped.PivotTable);
        spec.Note("A pivot table is three parts and a duplicated copy of the data: the cache " +
                  "definition describes the source fields, the cache records hold a snapshot of " +
                  "the rows, and the table definition says how they are laid out. Counted once, " +
                  "as one construct, however many parts it took.");
        spec.Note("The rendered result is ordinary cells on the target sheet. Those import like " +
                  "any others — which is the right outcome: the reader keeps what the author saw " +
                  "and loses only the ability to re-pivot it.");

        corpus.Emit(spec, path => Book.Create(path, spec, book =>
        {
            Tab source = book.Sheet("Source");
            source.Text("A1", "region").Text("B1", "amount");
            string[] regions = ["north", "south", "north", "south"];
            for (int i = 0; i < regions.Length; i++)
            {
                source.Text($"A{i + 2}", regions[i]);
                source.Number($"B{i + 2}", (i + 1) * 100);
            }

            Tab report = book.Sheet("Report");
            report.Text("A1", "region").Text("B1", "sum of amount");
            report.Text("A2", "north").Number("B2", 400);
            report.Text("A3", "south").Number("B3", 600);
            report.Text("A4", "grand total").Number("B4", 1000);
            report.Text("D1", "the rendered cells are ordinary and must survive the pivot being dropped");

            PivotTableCacheDefinitionPart cache = book.Part.AddNewPart<PivotTableCacheDefinitionPart>();
            Book.WriteXml(cache, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                      r:id="rId1" recordCount="4">
                  <cacheSource type="worksheet">
                    <worksheetSource ref="A1:B5" sheet="Source"/>
                  </cacheSource>
                  <cacheFields count="2">
                    <cacheField name="region" numFmtId="0">
                      <sharedItems count="2"><s v="north"/><s v="south"/></sharedItems>
                    </cacheField>
                    <cacheField name="amount" numFmtId="0">
                      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"
                                   containsInteger="1" minValue="100" maxValue="400"/>
                    </cacheField>
                  </cacheFields>
                </pivotCacheDefinition>
                """);

            PivotTableCacheRecordsPart records = cache.AddNewPart<PivotTableCacheRecordsPart>("rId1");
            Book.WriteXml(records, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
                  <r><x v="0"/><n v="100"/></r>
                  <r><x v="1"/><n v="200"/></r>
                  <r><x v="0"/><n v="300"/></r>
                  <r><x v="1"/><n v="400"/></r>
                </pivotCacheRecords>
                """);

            PivotTablePart pivot = report.PartOf.AddNewPart<PivotTablePart>();
            pivot.AddPart(cache);
            Book.WriteXml(pivot, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                      name="PivotTable1" cacheId="1" dataCaption="Values"
                                      applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"
                                      applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1">
                  <location ref="A1:B4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
                  <pivotFields count="2">
                    <pivotField axis="axisRow" showAll="0">
                      <items count="3"><item x="0"/><item x="1"/><item t="default"/></items>
                    </pivotField>
                    <pivotField dataField="1" showAll="0"/>
                  </pivotFields>
                  <rowFields count="1"><field x="0"/></rowFields>
                  <rowItems count="3">
                    <i><x v="0"/></i>
                    <i><x v="1"/></i>
                    <i t="grand"><x/></i>
                  </rowItems>
                  <colItems count="1"><i/></colItems>
                  <dataFields count="1">
                    <dataField name="sum of amount" fld="1" baseField="0" baseItem="0"/>
                  </dataFields>
                </pivotTableDefinition>
                """);

            book.ExtraWorkbookChildren.Add(new PivotCaches(
                new PivotCache { CacheId = 1, Id = book.Part.GetIdOfPart(cache) }));
        }));
    }

    static FixtureSpec New(string file, string milestone, string covers) =>
        new() { File = file, Family = "document", Milestone = milestone, Covers = covers };
}
