using NPOI.SS.UserModel;

namespace XlsxToPdfConverter.Diy
{
    public static class PrintDataConverter
    {
        public static PaperOrientation GetOrientation(IPrintSetup input)
        {
            if (input.Landscape)
            {
                return PaperOrientation.Landscape;
            }
            else
            {
                return PaperOrientation.Portrait;
            }
        }

        public static PaperSize GetPaperSize(IPrintSetup input)
        {
            if (
                input.PaperSize == (int)NPOI.SS.UserModel.PaperSize.A4 ||
                input.PaperSize == (int)NPOI.SS.UserModel.PaperSize.A4_Small)
            {
                return PaperSize.A4;
            }
            else if (input.PaperSize == (int)NPOI.SS.UserModel.PaperSize.A3)
            {
                return PaperSize.A3;
            }
            return PaperSize.A4;
        }

        public static (double top, double left, double right, double bottom) GetMargin(ISheet input)
        {
            // дюймы -> сантиметры -> пиксели
            const double magicNumber = 2.54 * 72;
            return (magicNumber * (input.GetMargin(MarginType.TopMargin) + input.GetMargin(MarginType.HeaderMargin)),
                magicNumber * input.GetMargin(MarginType.LeftMargin),
                magicNumber * input.GetMargin(MarginType.RightMargin),
                magicNumber * (input.GetMargin(MarginType.BottomMargin) + input.GetMargin(MarginType.FooterMargin)));
        }

        public static (double top, double bottom) GetHeaderAndFooterMargin(ISheet input)
        {
            // дюймы -> сантиметры -> пиксели
            const double magicNumber = 2.54 * 72;
            return (magicNumber * input.GetMargin(MarginType.HeaderMargin),
                magicNumber * input.GetMargin(MarginType.FooterMargin));
        }
    }
}