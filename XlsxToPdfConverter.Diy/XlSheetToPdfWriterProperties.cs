namespace XlsxToPdfConverter.Diy
{
    internal class XlSheetToPdfWriterProperties
    {
        public int ColNum { get; set; }

        public PaperSize PagePrintSize { get; set; }

        public PaperOrientation PaperOrientation { get; set; }

        public (double top, double left, double right, double bottom) PagePrintMargins { get; set; }

        public (double header, double footer) PageHeaderAndFooterMargins { get; set; }
    }
}
