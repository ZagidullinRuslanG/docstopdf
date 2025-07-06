namespace XlsxToPdfConverter.Diy
{
    public class PaperSize
    {
        private PaperSize(double width, double height)
        {
            Width = width;
            Heigth = height;
        }

        public double Width { get; }

        public double Heigth { get; }

        // точек в сантиметре
        private const double PPCm = 72;

        private const double MmToInch = 25.4;

        public static PaperSize A4 = new PaperSize(
            PPCm * 210 / MmToInch,
            PPCm * 297 / MmToInch);

        public static PaperSize A3 = new PaperSize(
            PPCm * 297 / MmToInch,
            PPCm * 420 / MmToInch);


        public static PaperSize A2 = new PaperSize(
            PPCm * 420 / MmToInch,
            PPCm * 594 / MmToInch);
    }
}
