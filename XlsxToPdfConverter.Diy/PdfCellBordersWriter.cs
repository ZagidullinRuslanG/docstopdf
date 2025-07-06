using System.Collections.Generic;
using NPOI.SS.UserModel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace XlsxToPdfConverter.Diy
{
    internal class PdfCellBordersWriter
    {
        private readonly PdfPage page;
        private readonly double printScale;
        private readonly Dictionary<(BorderStyle xlBorderStyle, short xlBorderColor), XPen> pens;

        public PdfCellBordersWriter(PdfPage page, double printScale)
        {
            this.page = page;
            this.printScale = printScale;
            pens = new Dictionary<(BorderStyle xlBorderStyle, short xlBorderColor), XPen>();
        }

        public void DrawBorders(XRect area, ICellStyle xlStyle, bool ignoreLeftAndTop)
        {
            using (var gfx = XGraphics.FromPdfPage(page))
            {
                if (xlStyle.BorderTop != BorderStyle.None && !ignoreLeftAndTop)
                {
                    DrawLine(gfx, xlStyle.BorderTop, xlStyle.TopBorderColor, area.TopLeft, area.TopRight);
                }
                if (xlStyle.BorderBottom != BorderStyle.None)
                {
                    DrawLine(gfx, xlStyle.BorderBottom, xlStyle.BottomBorderColor, area.BottomLeft, area.BottomRight);
                }
                if (xlStyle.BorderLeft != BorderStyle.None && !ignoreLeftAndTop)
                {
                    DrawLine(gfx, xlStyle.BorderLeft, xlStyle.LeftBorderColor, area.TopLeft, area.BottomLeft);
                }
                if (xlStyle.BorderRight != BorderStyle.None)
                {
                    DrawLine(gfx, xlStyle.BorderRight, xlStyle.RightBorderColor, area.TopRight, area.BottomRight);
                }
            }
        }

        private void DrawLine(XGraphics gfx, BorderStyle xlBorderStyle, short xlBorderColor, XPoint point1, XPoint point2)
        {
            XPen pen = GetPen(xlBorderStyle, xlBorderColor);

            double halfWidth = pen.Width / 2;

            double x1;
            double x2;
            if (point1.X < point2.X)
            {
                x1 = point1.X - halfWidth;
                x2 = point2.X + halfWidth;
            }
            else if (point1.X > point2.X)
            {
                x1 = point1.X + halfWidth;
                x2 = point2.X - halfWidth;
            }
            else
            {
                x1 = point1.X;
                x2 = point2.X;
            }
            double y1;
            double y2;
            if (point1.Y < point2.Y)
            {
                y1 = point1.Y - halfWidth;
                y2 = point2.Y + halfWidth;
            }
            else if (point1.Y > point2.Y)
            {
                y1 = point1.Y + halfWidth;
                y2 = point2.Y - halfWidth;
            }
            else
            {
                y1 = point1.Y;
                y2 = point2.Y;
            }

            gfx.DrawLine(pen, x1, y1, x2, y2);
        }


        private XPen GetPen(BorderStyle xlBorderStyle, short xlBorderColor)
        {
            if (!pens.TryGetValue((xlBorderStyle, xlBorderColor), out XPen pen))
            {
                if (!xlBorderStyleWidthMap.TryGetValue(xlBorderStyle, out double width))
                {
                    width = defaultBorderWidth;
                }
                width /= printScale;

                pen = new XPen(XColor.FromKnownColor(XKnownColor.Black), width);
                pens.Add((xlBorderStyle, xlBorderColor), pen);
            }
            return pen;
        }

        private static readonly double defaultBorderWidth = 0.25;
        private static readonly Dictionary<BorderStyle, double> xlBorderStyleWidthMap =
            new Dictionary<BorderStyle, double>()
            {
                { BorderStyle.Thin, 0.25 },
                { BorderStyle.Medium, 0.5 },
                { BorderStyle.Thick, 0.75 },
            };
    }
}
