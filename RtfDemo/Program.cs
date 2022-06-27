using Elistia.DotNetRtfWriter;
using System.Diagnostics;

namespace RtfWriter.Demo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // Create document by specifying paper size and orientation,
            // and default language.
            var doc = new RtfDocument(PaperSize.A4, PaperOrientation.Landscape, Lcid.English);

            // Create fonts and colors for later use
            var times = doc.CreateFont("Times New Roman");
            var courier = doc.CreateFont("Courier New");
            var red = doc.CreateColor(new RtfColor("ff0000"));
            var blue = doc.CreateColor(new RtfColor(0, 0, 255));
            var white = doc.CreateColor(new RtfColor(255, 255, 255));
            var colorTableHeader = doc.CreateColor(new RtfColor("76923C"));
            var colorTableRow = doc.CreateColor(new RtfColor("D6E3BC"));
            var colorTableRowAlt = doc.CreateColor(new RtfColor("FFFFFF"));

            // Don't instantiate RtfTable, RtfParagraph, and RtfImage objects by using
            // ``new'' keyword. Instead, use add* method in objects derived from
            // RtfBlockList class. (See Demos.)
            RtfTable table;
            RtfParagraph par;
            RtfImage img;
            // Don't instantiate RtfCharFormat by using ``new'' keyword, either.
            // An addCharFormat method are provided by RtfParagraph objects.
            RtfCharFormat fmt;


            // ==========================================================================
            // Demo 1: Font Setting
            // ==========================================================================
            // If you want to use Latin characters only, it is as simple as assigning
            // ``Font'' property of RtfCharFormat objects. If you want to render Far East
            // characters with some font, and Latin characters with another, you may
            // assign the Far East font to ``Font'' property and the Latin font to
            // ``AnsiFont'' property.
            par = doc.AddParagraph();
            par.Alignment = Align.Left;
            par.DefaultCharFormat.Font = times;
            par.DefaultCharFormat.AnsiFont = courier;
            par.SetText("Testing\n");


            // ==========================================================================
            // Demo 2: Character Formatting
            // ==========================================================================
            par = doc.AddParagraph();
            par.DefaultCharFormat.Font = times;
            par.SetText("Demo2: Character Formatting");
            // Besides setting default character formats of a paragraph, you can specify
            // a range of characters to which formatting is applied. For convenience,
            // let's call it range formatting. The following section sets formatting
            // for the 4th, 5th, ..., 8th characters in the paragraph. (Note: the first
            // character has an index of 0)
            fmt = par.AddCharFormat(4, 8);
            fmt.FgColor = blue;
            fmt.BgColor = red;
            fmt.FontSize = 18;
            // Sets another range formatting. Note that when range formatting overlaps,
            // the latter formatting will overwrite the former ones. In the following,
            // formatting for the 8th chacacter is overwritten.
            fmt = par.AddCharFormat(8, 10);
            fmt.FontStyle.AddStyle(FontStyleFlag.Bold);
            fmt.FontStyle.AddStyle(FontStyleFlag.Underline);
            fmt.Font = courier;


            // ==========================================================================
            // Demo 3: Footnote
            // ==========================================================================
            par = doc.AddParagraph();
            par.SetText("Demo3: Footnote");
            // In this example, the footnote is inserted just after the 7th character in
            // the paragraph.
            par.AddFootnote(7).AddParagraph().SetText("Footnote details here.");


            // ==========================================================================
            // Demo 4: Header and Footer
            // ==========================================================================
            // You may use ``Header'' and ``Footer'' properties of RtfDocument objects to
            // specify information to be displayed in the header and footer of every page,
            // respectively.
            par = doc.Footer.AddParagraph();
            par.SetText("Demo4: Page: / Date: Time:");
            par.Alignment = Align.Center;
            par.DefaultCharFormat.FontSize = 15;
            // You may insert control words, including page number, total pages, date and
            // time, into the header and/or the footer.
            par.AddControlWord(12, RtfFieldControlWord.FieldType.Page);
            par.AddControlWord(13, RtfFieldControlWord.FieldType.NumPages);
            par.AddControlWord(19, RtfFieldControlWord.FieldType.Date);
            par.AddControlWord(25, RtfFieldControlWord.FieldType.Time);
            // Here we also add some text in header.
            par = doc.Header.AddParagraph();
            par.SetText("Demo4: Header");


            // ==========================================================================
            // Demo 5: Image
            // ==========================================================================
            img = doc.AddImage("../../demo5.jpg", RtfImageType.Jpg);
            // You may set the width only, and let the height be automatically adjusted
            // to keep aspect ratio.
            img.Width = 130;
            // Place the image on a new page. The ``StartNewPage'' property is also supported
            // by paragraphs and tables.
            //img.StartNewPage = true;
            img.StartNewPara = true;


            // ==========================================================================
            // demo 6: 表格
            // ==========================================================================
            // Please be careful when dealing with tables, as most crashes come from them.
            // If you follow steps below, the resulting RTF is not likely to crash your
            // MS Word.
            // 
            // Step 1. Plan and draw the table you want on a scratch paper.
            // Step 2. Start with a MxN regular table.
            table = doc.AddTable(5, 4, 415.2f, 12);
            table.Margins[Direction.Bottom] = 20;
            table.SetInnerBorder(BorderStyle.Dotted, 1f);
            table.SetOuterBorder(BorderStyle.Single, 2f);

            table.HeaderBackgroundColor = colorTableHeader;
            table.RowBackgroundColor = colorTableRow;
            table.RowAltBackgroundColor = colorTableRowAlt;


            // Step 3. (Optional) Set text alignment for each cell, row height, column width,
            //			border style, etc.
            for (var i = 0; i < table.RowCount; i++) {
                for (var j = 0; j < table.ColCount; j++) {
                    table.Cell(i, j).AddParagraph().SetText("CELL " + i.ToString() + "," + j.ToString());
                }
            }

            // Step 4. Merge cells so that the resulting table would look like the one you drew
            //			on paper. One cell cannot be merged twice. In this way, we can construct
            //			almost all kinds of tables we need.
            table.Merge(1, 0, 3, 1);
            // Step 5. You may start inserting content for each cell. Actually, it is adviced
            //			that the only thing you do after merging cell is inserting content.
            table.Cell(4, 3).BackgroundColor = red;
            table.Cell(4, 3).AddParagraph().SetText("Demo6: Table");


            // ==========================================================================
            // Demo 7: ``Two in one'' format
            // ==========================================================================
            // This format is provisioned for Far East languages. This demo uses Traditional
            // Chinese as an example.
            par = doc.AddParagraph();
            par.SetText("Demo7: aaa並排文字aaa");
            fmt = par.AddCharFormat(10, 13);
            fmt.TwoInOneStyle = TwoInOneStyle.Braces;
            fmt.FontSize = 16;


            // ==========================================================================
            // Demo 7.1: Hyperlink
            // ==========================================================================
            par = doc.AddParagraph();
            par.SetText("Demo 7.1: Hyperlink to target (Demo9)");
            fmt = par.AddCharFormat(10, 18);
            fmt.LocalHyperlink = "target";
            fmt.LocalHyperlinkTip = "Link to target";
            fmt.FgColor = blue;


            // ==========================================================================
            // Demo 8: New page
            // ==========================================================================
            par = doc.AddParagraph();
            par.StartNewPage = true;
            par.SetText("Demo8: New page");


            // ==========================================================================
            // Demo 9: Set bookmark
            // ==========================================================================
            par = doc.AddParagraph();
            par.SetText("Demo9: Set bookmark");
            fmt = par.AddCharFormat(0, 18);
            fmt.Bookmark = "target";


            // ==========================================================================
            // Save
            // ==========================================================================
            // You may also retrieve RTF code string by calling to render() method of
            // RtfDocument objects.
            doc.Save("Demo.rtf");


            // ==========================================================================
            // Open the RTF file we just saved
            // ==========================================================================
            var p = new Process { StartInfo = { FileName = "Demo.rtf" } };
            p.Start();
        }
    }
}
