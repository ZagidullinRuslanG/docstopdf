using System;
using System.Collections.Generic;
using NpoiHelpers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Запись текста в заданную область с учетом выравнивания, шрифта, отступов и других параметров.
    /// </summary>
    internal class PdfCellTextWriter
    {
        public void Write(
            PdfPage page,
            XRect area,
            double padding,
            ICellStyle xlCellStyle,
            IFont xlFont,
            string text)
        {
            var font = new PdfCellFontMapper().Get(xlFont);

            var brush = XBrushes.Black;
            var xssfColor = ((XSSFFont)xlFont)?.GetXSSFColor();
            if (xssfColor != null)
            {
                try
                {
                    var color = XssfColorConverter.ConvertToRgb(xssfColor);
                    brush = new XSolidBrush(XColor.FromArgb(color[0], color[1], color[2]));
                }
                catch
                {
                }
            }

            var paragraphAlignment = GetParagraphAlignment(xlCellStyle);

            Write(
                page,
                area,
                padding,
                font,
                brush,
                paragraphAlignment,
                xlCellStyle.VerticalAlignment == VerticalAlignment.Center,
                text);
        }

        public void Write(
            PdfPage page,
            XRect area,
            double padding,
            XFont font,
            XBrush brush,
            XParagraphAlignment paragraphAlignment,
            bool isVerticalCentered,
            string text)
        {
            area.X += padding;
            area.Y += padding;
            area.Height = Math.Max(0.0, area.Height - (2 * padding));
            area.Width = Math.Max(0.0, area.Width - (2 * padding));

            using (var gfx = XGraphics.FromPdfPage(page))
            {
                // Определяем необходимую высоту, чтобы сделать вертикальное выравнивание
                // (XTextFormatter поддерживает только XStringFormat.TopLeft).
                var tf2 = new XTextFormatterEx2(gfx);
                tf2.Alignment = paragraphAlignment;
                tf2.PrepareDrawString(text, font, brush, area, out int lastFittingChar, out double neededHeight);

                if (isVerticalCentered)
                {
                    // Симметрично убираем лишнюю высоту, чтобы получить эффект вертикального выравнивания по центру
                    // относительно исходной области area.
                    double excessHeight = area.Height - neededHeight;
                    area.Y += excessHeight / 2;
                    area.Height -= excessHeight;
                }

                tf2.DrawString();
            }
        }

        public void Write(
            PdfPage page,
            XRect area,
            double padding,
            ICellStyle xlCellStyle,
            StyleList styles,
            string text)
        {
            var paragraphAlignment = GetParagraphAlignment(xlCellStyle);
            area.X += padding;
            area.Y += padding;
            area.Height = Math.Max(0.0, area.Height - (2 * padding));
            area.Width = Math.Max(0.0, area.Width - (2 * padding));

            List<StyleTextPair> content = styles.SplitAsApplied(text);
            var offsets = new List<double>();

            using (var gfx = XGraphics.FromPdfPage(page))
            {
                // Определяем высоту абзацев, чтобы сделать вертикальное выравнивание
                var tf2 = new XTextFormatterEx2(gfx);
                tf2.Alignment = paragraphAlignment;
                tf2.LayoutRectangle = area;

                double neededHeight = 0;

                foreach (var item in content)
                {
                    XColor xColor = XColor.FromArgb(0, 0, 0);
                    var xssfColor = ((XSSFFont)item.Font)?.GetXSSFColor();
                    if (xssfColor != null)
                    {
                        try
                        {
                            var rgbColor = XssfColorConverter.ConvertToRgb(xssfColor);
                            xColor = XColor.FromArgb(rgbColor[0], rgbColor[1], rgbColor[2]);
                        }
                        catch
                        {
                        }
                    }
                    tf2.AppendPreparedDrawString(item.Text, new PdfCellFontMapper().Get(item.Font), new XSolidBrush(xColor), out int lastFittingChar, out neededHeight);
                    offsets.Add(neededHeight);
                }

                if (xlCellStyle.VerticalAlignment == VerticalAlignment.Center)
                {
                    // Симметрично убираем лишнюю высоту, чтобы получить эффект вертикального выравнивания по центру
                    // относительно исходной области area.
                    double excessHeight = area.Height - neededHeight;
                    area.Y += excessHeight / 2;
                    area.Height -= excessHeight;
                    tf2.LayoutRectangle = area;
                }

                tf2.DrawString();
            }
        }

        private XParagraphAlignment GetParagraphAlignment(ICellStyle xlCellStyle)
        {
            switch (xlCellStyle.Alignment)
            {
                case HorizontalAlignment.General:
                case HorizontalAlignment.Left:
                    return XParagraphAlignment.Left;
                case HorizontalAlignment.Center:
                    return XParagraphAlignment.Center;
                case HorizontalAlignment.Right:
                    return XParagraphAlignment.Right;
                default:
                    throw new NotSupportedException();
            }
        }
    }
}
