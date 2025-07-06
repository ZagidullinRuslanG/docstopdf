using System;
using System.Collections.Generic;
using System.Linq;
using NpoiHelpers;
using NPOI.SS.UserModel;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

namespace XlsxToPdfConverter.Diy
{
    internal class PdfWriter : IDisposable
    {
        private readonly PdfDocument doc;

        private PdfPartitionLayout partLayout;
        private List<PageInfo> pages = new ();

        static PdfWriter()
        {
            GlobalFontSettings.FontResolver = new CustomFontResolver();
        }

        public PdfWriter()
        {
            doc = new PdfDocument();

            doc.SecuritySettings.PermitAnnotations = false;
            doc.SecuritySettings.PermitAssembleDocument = false;
            doc.SecuritySettings.PermitExtractContent = false;
            doc.SecuritySettings.PermitFormsFill = false;
            doc.SecuritySettings.PermitFullQualityPrint = true;
            doc.SecuritySettings.PermitModifyDocument = false;
            doc.SecuritySettings.PermitPrint = true;
        }

        public void AddPartition(PdfPartitionLayout layout)
        {
            partLayout = layout;
            pages = new ();

            var curPage = doc.AddPage();
            curPage.Width = new XUnit(layout.PageActualSize.Width);
            curPage.Height = new XUnit(layout.PageActualSize.Height);
            var curPageCellBordersWriter = new PdfCellBordersWriter(curPage, partLayout.PrintScale);
            var curPageCellBackgroundWriter = new PdfCellBackgroundWriter(curPage, partLayout.PrintScale);

            pages.Add(new ()
            {
                PrevPartPagesSumXlHeight = 0,
                Page = curPage,
                BordersWriter = curPageCellBordersWriter,
                BackgroundWriter = curPageCellBackgroundWriter
            });
        }

        /// <summary>
        /// Пишет на указанном листе в указанной позиции колонтитула текст.
        /// </summary>
        public void WriteHeaderOrFooter(
            int pageNum,
            HeaderOrFooterPosition positions,
            string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            bool isHeader = (positions & HeaderOrFooterPosition.Top) != 0;
            PdfPage page = doc.Pages[pageNum];
            double y, x, height;
            XParagraphAlignment align;

            if ((positions & HeaderOrFooterPosition.Left) != 0)
            {
                x = partLayout.PageActualMargins.Left;
                align = XParagraphAlignment.Left;
            }
            else if ((positions & HeaderOrFooterPosition.Right) != 0)
            {
                x = partLayout.PageActualMargins.Left + (partLayout.PageActualHeaderAndFooter.HorisontalSize * 2);
                align = XParagraphAlignment.Right;
            }
            //if (( positions & HeaderOrFooterPosition.center ) != 0)
            else
            {
                x = partLayout.PageActualMargins.Left + partLayout.PageActualHeaderAndFooter.HorisontalSize;
                align = XParagraphAlignment.Center;
            }

            if (isHeader)
            {
                y = partLayout.PageActualHeaderAndFooter.Header;
                height = partLayout.PageActualHeaderAndFooter.HeaderHeight;
            }
            else
            {
                y = partLayout.PageActualHeaderAndFooter.Footer;
                height = partLayout.PageActualHeaderAndFooter.FooterHeight;
            }


            XFont font = new XFont("Arial", 9, 0);

            new PdfCellTextWriter().Write(
                page,
                new XRect(x,
                    y,
                    partLayout.PageActualHeaderAndFooter.HorisontalSize,
                    height),
                0,
                font,
                XBrushes.Black,
                align,
                true,
                text);
        }

        public int GetCurrentPageNum()
        {
            return doc.PageCount - 1;
        }

        public void Save(string path)
        {
            doc.Save(path);
        }

        // Перед вызовом нужно проверять методом CheckAreaAndAddPageIfNeeded необходимость создания следующей страницы.
        private XRect GetArea(Rect xlArea)
        {
            var rect = new XRect(
                   xlArea.X + partLayout.PageActualMargins.Left,
                   xlArea.Y - GetPageInfo(xlArea.Y).PrevPartPagesSumXlHeight + partLayout.PageActualMargins.Top,
                   xlArea.W,
                   xlArea.H);

            return rect;
        }

        private void CheckAreaAndAddPageIfNeeded(Rect xlArea)
        {
            var lastPageInfo = pages.Last();

            // Если не помещается на последнюю страницу...
            if (xlArea.Y - lastPageInfo.PrevPartPagesSumXlHeight + xlArea.H >
                partLayout.PageActualSize.Height - partLayout.PageActualMargins.Top - partLayout.PageActualMargins.Bottom)
            {
                // ...создаем следующую

                var curPage = doc.AddPage();
                curPage.Width = partLayout.PageActualSize.Width;
                curPage.Height = partLayout.PageActualSize.Height;
                var curPageCellBordersWriter = new PdfCellBordersWriter(curPage, partLayout.PrintScale);
                var curPageCellBackgroundWriter = new PdfCellBackgroundWriter(curPage, partLayout.PrintScale);

                pages.Add(new ()
                {
                    PrevPartPagesSumXlHeight = xlArea.Y,
                    Page = curPage,
                    BordersWriter = curPageCellBordersWriter,
                    BackgroundWriter = curPageCellBackgroundWriter
                });
            }
        }

        public void WriteBorders(ICellStyle xlCellStyle, Rect xlArea, bool ignoreLeftAndTop = false)
        {
            CheckAreaAndAddPageIfNeeded(xlArea);
            XRect area = GetArea(xlArea);
            GetPageInfo(xlArea.Y).BordersWriter.DrawBorders(area, xlCellStyle, ignoreLeftAndTop);
        }

        public void WriteBackground(ICellStyle xlCellStyle, Rect xlArea, bool ignoreLeftAndTop = false)
        {
            CheckAreaAndAddPageIfNeeded(xlArea);
            XRect area = GetArea(xlArea);
            GetPageInfo(xlArea.Y).BackgroundWriter.DrawBackground(area, xlCellStyle, ignoreLeftAndTop);
        }

        public void WriteText(
            string text,
            ICellStyle xlCellStyle,
            IFont xlFont,
            Rect xlArea)
        {
            CheckAreaAndAddPageIfNeeded(xlArea);
            XRect area = GetArea(xlArea);

            double printPadding = 1;
            double actualPadding = printPadding * partLayout.PageActualSize.Width / partLayout.PagePrintSize.Width;

            new PdfCellTextWriter().Write(GetPageInfo(xlArea.Y).Page, area, actualPadding, xlCellStyle, xlFont, text);
        }

        public void WriteText(
            string text,
            ICellStyle xlCellStyle,
            StyleList styles,
            Rect xlArea)
        {
            CheckAreaAndAddPageIfNeeded(xlArea);
            XRect area = GetArea(xlArea);

            double printPadding = 1;
            double actualPadding = printPadding * partLayout.PageActualSize.Width / partLayout.PagePrintSize.Width;

            new PdfCellTextWriter().Write(GetPageInfo(xlArea.Y).Page, area, actualPadding, xlCellStyle, styles, text);
        }

        /// <summary>
        /// Задает область, которую нельзя разделять на разные страницы PDF-документа.
        /// После вызова этого метода возможен переход на следующую PDF-страницу,
        /// после которого запись в предыдущую будет невозможна.
        /// </summary>
        public void SetNotSplittableByPagesArea(Rect xlArea)
        {
            CheckAreaAndAddPageIfNeeded(xlArea);
        }

        public void Dispose()
        {
            if (doc != null)
            {
                doc.Dispose();
            }
        }

        private PageInfo GetPageInfo(double y) => pages.Last(p => p.PrevPartPagesSumXlHeight <= y);

        private class PageInfo
        {
            public double PrevPartPagesSumXlHeight { get; set; }

            public PdfPage Page { get; set; }

            public PdfCellBordersWriter BordersWriter { get; set; }

            public PdfCellBackgroundWriter BackgroundWriter { get; set; }
        }
    }
}
