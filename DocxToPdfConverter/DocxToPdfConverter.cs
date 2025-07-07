using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Linq;

namespace DocxToPdfConverter
{
    public class DocxToPdfConverter
    {
        public void Convert(string docxPath, string pdfPath)
        {
            using (PdfDocument pdf = new PdfDocument())
            {
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                double y = 40;
                double left = 40;
                double right = page.Width - 40;
                double lineHeight = 18;
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var numbering = wordDoc.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
                    foreach (var element in body.Elements())
                    {
                        if (element is Paragraph para)
                        {
                            string styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
                            XFont font = GetFontForStyle(styleId);
                            double curLineHeight = GetLineHeightForStyle(styleId);
                            double x = left;
                            // Списки
                            var numProp = para.ParagraphProperties?.NumberingProperties;
                            if (numProp != null)
                            {
                                string marker = GetListMarker(numProp, numbering);
                                if (!string.IsNullOrEmpty(marker))
                                {
                                    gfx.DrawString(marker, font, XBrushes.Black, new XRect(x, y, 30, curLineHeight), XStringFormats.TopLeft);
                                    x += gfx.MeasureString(marker, font).Width + 5;
                                }
                            }
                            foreach (var run in para.Elements<Run>())
                            {
                                string text = string.Concat(run.Elements<Text>().Select(t => t.Text));
                                if (string.IsNullOrEmpty(text)) continue;
                                var runProps = run.RunProperties;
                                var runFont = GetFontForRun(font, runProps);
                                gfx.DrawString(text, runFont, XBrushes.Black, new XRect(x, y, right - left, curLineHeight), XStringFormats.TopLeft);
                                x += gfx.MeasureString(text, runFont).Width;
                            }
                            y += curLineHeight;
                            if (y > page.Height - 40)
                            {
                                page = pdf.AddPage();
                                gfx.Dispose();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 40;
                            }
                        }
                        else if (element is Table table)
                        {
                            y = DrawTable(table, gfx, left, y, right - left);
                            if (y > page.Height - 40)
                            {
                                page = pdf.AddPage();
                                gfx.Dispose();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 40;
                            }
                        }
                    }
                }
                gfx.Dispose();
                pdf.Save(pdfPath);
            }
        }

        private string GetListMarker(NumberingProperties numProp, Numbering numbering)
        {
            if (numProp.NumberingId == null) return null;
            string marker = "• ";
            if (numbering != null && numProp.NumberingId != null)
            {
                var numId = numProp.NumberingId.Val.Value;
                var abstractNumId = numbering.Elements<NumberingInstance>()
                    .FirstOrDefault(n => n.NumberID.Value == numId)?
                    .AbstractNumId?.Val?.Value;
                if (abstractNumId != null)
                {
                    var abstractNum = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.AbstractNum>()
                        .FirstOrDefault(n => n.AbstractNumberId.Value == abstractNumId);
                    if (abstractNum != null)
                    {
                        var lvl = abstractNum.Elements<Level>().FirstOrDefault();
                        if (lvl != null && lvl.NumberingFormat != null)
                        {
                            if (lvl.NumberingFormat.Val == NumberFormatValues.Bullet)
                                marker = "• ";
                            else if (lvl.NumberingFormat.Val == NumberFormatValues.Decimal)
                                marker = "1. ";
                        }
                    }
                }
            }
            return marker;
        }

        private double DrawTable(Table table, XGraphics gfx, double left, double y, double tableWidth)
        {
            double cellPadding = 4;
            int columns = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 1;
            double cellWidth = tableWidth / columns;
            double curY = y;
            XFont font = new XFont("Arial", 12, XFontStyleEx.Regular);
            foreach (var row in table.Elements<TableRow>())
            {
                double x = left;
                double rowHeight = 18;
                foreach (var cell in row.Elements<TableCell>())
                {
                    string cellText = string.Join(" ", cell.Descendants<Text>().Select(t => t.Text));
                    gfx.DrawRectangle(XPens.Black, x, curY, cellWidth, rowHeight);
                    gfx.DrawString(cellText, font, XBrushes.Black, new XRect(x + cellPadding, curY + cellPadding, cellWidth - 2 * cellPadding, rowHeight - 2 * cellPadding), XStringFormats.TopLeft);
                    x += cellWidth;
                }
                curY += rowHeight;
            }
            return curY + 8;
        }

        private XFont GetFontForStyle(string styleId)
        {
            if (styleId == "Heading1") return new XFont("Arial", 20, XFontStyleEx.Bold);
            if (styleId == "Heading2") return new XFont("Arial", 16, XFontStyleEx.Bold);
            if (styleId == "Heading3") return new XFont("Arial", 14, XFontStyleEx.Bold);
            return new XFont("Arial", 12, XFontStyleEx.Regular);
        }
        private double GetLineHeightForStyle(string styleId)
        {
            if (styleId == "Heading1") return 28;
            if (styleId == "Heading2") return 22;
            if (styleId == "Heading3") return 18;
            return 16;
        }
        private XFont GetFontForRun(XFont baseFont, RunProperties runProps)
        {
            XFontStyleEx style = XFontStyleEx.Regular;
            if (runProps != null)
            {
                if (runProps.Bold != null) style |= XFontStyleEx.Bold;
                if (runProps.Italic != null) style |= XFontStyleEx.Italic;
                if (runProps.Underline != null) style |= XFontStyleEx.Underline;
            }
            return new XFont(baseFont.FontFamily.Name, baseFont.Size, style);
        }
    }
} 