using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiHelpers
{
    public static class CellStyleExtensions
    {
        /// <summary>
        /// Установить сплошной цвет RGB у заднего фона в стиле ячейки.
        /// </summary>
        /// <param name="style">Стиль ячейки.</param>
        /// <param name="red">Красный.</param>
        /// <param name="green">Зеленый.</param>
        /// <param name="blue">Синий.</param>
        public static void SetFillColorRgbSolid(this ICellStyle style, byte red, byte green, byte blue)
            => SetFillXssfColor(style, new NPOI.XSSF.UserModel.XSSFColor(new[] { red, green, blue }));

        /// <summary>
        /// Установить сплошной системный цвет у заднего фона в стиле ячейки.
        /// </summary>
        /// <param name="style">Стиль ячейки.</param>
        /// <param name="color">Цвет.</param>
        public static void SetFillColorSystemSolid(this ICellStyle style, System.Drawing.Color color) =>
            SetFillXssfColor(style, new NPOI.XSSF.UserModel.XSSFColor(new[] { color.R, color.G, color.B }));

        /// <summary>
        /// Установить сплошной цвет XSSFColor в стиле ячейке.
        /// </summary>
        /// <param name="style">Стиль ячейки.</param>
        /// <param name="color">Цвет.</param>
        private static void SetFillXssfColor(ICellStyle style, NPOI.XSSF.UserModel.XSSFColor color)
        {
            ((NPOI.XSSF.UserModel.XSSFCellStyle)style).SetFillForegroundColor(color);
            ((NPOI.XSSF.UserModel.XSSFCellStyle)style).SetFillBackgroundColor(color);
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}