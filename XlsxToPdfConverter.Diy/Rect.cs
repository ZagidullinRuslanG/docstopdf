namespace XlsxToPdfConverter.Diy
{
    internal class Rect
    {
        public Rect()
        {
        }

        public Rect(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            W = width;
            H = height;
        }

        public double X { get; set; }

        public double Y { get; set; }

        public double W { get; set; }

        public double H { get; set; }
    }
}
