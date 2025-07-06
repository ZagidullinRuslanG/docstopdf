namespace XlsxToPdfConverter.Diy
{
    internal class PdfPartitionLayout
    {
        public PdfPartitionLayout(
            double xlActualWidth,
            PaperSize pagePrintSize,
            PaperOrientation paperOrientation,
            (double top, double left, double right, double bottom) pagePrintMargins,
            (double header, double footer) pageHeaderAndFooter)
        {
            PagePrintMargins = pagePrintMargins;
            if (paperOrientation == PaperOrientation.Portrait)
            {
                PagePrintSize = (pagePrintSize.Width, pagePrintSize.Heigth);
            }
            else
            {
                PagePrintSize = (pagePrintSize.Heigth, pagePrintSize.Width);
            }

            PrintScale =
                (PagePrintSize.Width - PagePrintMargins.Left - PagePrintMargins.Right) /
                xlActualWidth;

            PageActualMargins = (
                PagePrintMargins.Top * PrintScale,
                PagePrintMargins.Left * PrintScale,
                PagePrintMargins.Right * PrintScale,
                PagePrintMargins.Bottom * PrintScale);

            double pageRatio = PagePrintSize.Height / PagePrintSize.Width;
            double pageActualWidth = xlActualWidth + PageActualMargins.Left + PageActualMargins.Right;
            PageActualSize = (pageActualWidth, pageActualWidth * pageRatio);

            PageActualHeaderAndFooter = (
                (PagePrintMargins.Top - pageHeaderAndFooter.header) * PrintScale,
                PageActualSize.Height - PageActualMargins.Bottom,
                xlActualWidth / 3,
                pageHeaderAndFooter.header * PrintScale,
                pageHeaderAndFooter.footer * PrintScale);
        }

        public double PrintScale { get; }

        public (double Top, double Left, double Right, double Bottom) PagePrintMargins { get; }

        public (double Top, double Left, double Right, double Bottom) PageActualMargins { get; }

        public (double Width, double Height) PagePrintSize { get; }

        public (double Width, double Height) PageActualSize { get; set; }

        public (double Header, double Footer, double HorisontalSize, double HeaderHeight, double FooterHeight) PageActualHeaderAndFooter { get; }
    }
}
