using System;
using NPOI.SS.UserModel;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;

namespace XlsxToPdfConverter.Diy
{
    internal class PdfСellAlignMapper
    {
        public XStringFormat Convert(ICellStyle xlStyle)
        {
            XStringAlignment xStringAlignment;
            switch (xlStyle.Alignment)
            {
                case HorizontalAlignment.General:
                case HorizontalAlignment.Left:
                    xStringAlignment = XStringAlignment.Near;
                    break;
                case HorizontalAlignment.Center:
                    xStringAlignment = XStringAlignment.Center;
                    break;
                case HorizontalAlignment.Right:
                    xStringAlignment = XStringAlignment.Far;
                    break;
                default:
                    throw new NotImplementedException();
            }

            XLineAlignment xLineAlignment;
            switch (xlStyle.VerticalAlignment)
            {
                case VerticalAlignment.Top:
                    xLineAlignment = XLineAlignment.Near;
                    break;
                case VerticalAlignment.Center:
                    xLineAlignment = XLineAlignment.Center;
                    break;
                case VerticalAlignment.Bottom:
                    xLineAlignment = XLineAlignment.Far;
                    break;
                default:
                    throw new NotImplementedException();
            }

            return new XStringFormat()
            {
                Alignment = xStringAlignment,
                LineAlignment = xLineAlignment,
            };
        }
    }
}
