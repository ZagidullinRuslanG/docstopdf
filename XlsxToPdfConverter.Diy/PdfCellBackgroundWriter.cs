using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace XlsxToPdfConverter.Diy
{
    internal class PdfCellBackgroundWriter
    {
        private static readonly double defaultBorderWidth = 0.25;
        private static readonly Dictionary<BorderStyle, double> xlBorderStyleWidthMap =
            new Dictionary<BorderStyle, double>()
            {
                { BorderStyle.None, 0 },
                { BorderStyle.Thin, 0.25 },
                { BorderStyle.Medium, 0.5 },
                { BorderStyle.Thick, 0.75 },
            };

        private readonly PdfPage page;
        private readonly double printScale;

        public PdfCellBackgroundWriter(PdfPage page, double printScale)
        {
            this.page = page;
            this.printScale = printScale;
        }

        public void DrawBackground(XRect area, ICellStyle xlStyle, bool ignoreLeftAndTop)
        {
            using (var gfx = XGraphics.FromPdfPage(page))
            {
                var top = area.Top;
                var left = area.Left;
                if (!ignoreLeftAndTop)
                {
                    top += xlBorderStyleWidthMap.GetValueOrDefault(xlStyle.BorderTop, defaultBorderWidth);
                    left += xlBorderStyleWidthMap.GetValueOrDefault(xlStyle.BorderLeft, defaultBorderWidth);
                }

                var bottom =
                    area.Bottom - xlBorderStyleWidthMap.GetValueOrDefault(xlStyle.BorderBottom, defaultBorderWidth);
                var right =
                    area.Right - xlBorderStyleWidthMap.GetValueOrDefault(xlStyle.BorderRight, defaultBorderWidth);

                if (top < bottom && left < right)
                {
                    DrawRect(gfx, xlStyle, new XRect(left, top, right - left, bottom - top));
                }
            }
        }

        private void DrawRect(XGraphics gfx, ICellStyle xlStyle, XRect area)
        {
            var brush = GetBrush(xlStyle);
            if (brush is not null)
            {
                gfx.DrawRectangle(brush, area);
            }
        }

        private XBrush GetBrush(ICellStyle xlStyle)
        {
            // После локальной правки шаблона отчета в ячейках с заливкой (настроенных в шаблоне) FillForegroundColorColor может быть null.
            // Поэтому такие ячейки после конвертации в PDF, выводятся без заливки (предположительно - стандартные заливки).
            // Возможный вариант лечения: пересохранить проблемные ячейки с чуть измененным элементом RGB (+-1).
            var color = ((XSSFColor)xlStyle.FillForegroundColorColor)?.RGBWithTint;
            if (xlStyle.FillPattern == FillPattern.SolidForeground && color != null)
            {
                var xColor = XColor.FromArgb(color[0], color[1], color[2]);
                return new XSolidBrush(xColor);
            }

            return null;
        }
    }
}