using XlsxToPdfConverter.Diy;
using NPOI.SS.UserModel;
using System;
using System.Globalization;

namespace XlsxToPdfConverter.Diy
{
    internal static class XlCellEx
    {
        private static readonly ILogger logger = new ConsoleLogger();

        public static bool IsEmpty(this ICell cell) =>
            (cell.CellType == CellType.Blank) ||
            (cell.CellType == CellType.String && string.IsNullOrWhiteSpace(cell.StringCellValue));

        public static bool IsNumber(this ICell cell) =>
            cell.CellType == CellType.Numeric &&
            !DateUtil.IsCellDateFormatted(cell);

        public static string AsText(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Blank:
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // попадаем сюда по условию выше, что ячейка в формате даты и не пустая, поэтому берем Value
                        DateTime date = cell.DateCellValue;
                        ICellStyle style = cell.CellStyle;

                        // note: Формат строки форматирования Excel отличается от C#.
                        string xlFormat = style.GetDataFormatString();
                        string format;
                        if (xlFormat == @"dd/mm/yy\ h:mm;@")
                        {
                            format = "dd.MM.yy HH:mm";
                        }
                        // если при выводе даты в Excel явно не преобразовывать DateTime в строку
                        // с форматом даты [dateTime.ToString("dd.MM.yyyy HH:mm")],
                        // то GetDataFormatString() возвращает @"dd/mm/yy;@"
                        else if (xlFormat == @"m/d/yy" || xlFormat == @"dd/mm/yy;@")
                        {
                            format = "dd.MM.yy";
                        }
                        else
                        {
                            const string defaultFormat = "dd.MM.yy";
                            logger.Warn($"Неподдерживаемое значение формата отображения даты из XLSX ({xlFormat}), " +
                                $"используется форматирование по умолчанию ({defaultFormat}).");
                            format = defaultFormat;
                        }
                        return date.ToString(format, CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        string newText;

                        // 2 -> id, соответствующий 2 десятичным знакам после запятой в настройке формата ячейки в Excel
                        // Не всегда совпадает с фактическим значением знаков!
                        // Например, 166 -> 3 десятичных знака
                        if (cell.CellStyle.DataFormat == 2)
                        {
                            newText = Math.Round(cell.NumericCellValue, 2, MidpointRounding.AwayFromZero).ToString();
                        }
                        // 9 -> процентный формат числа
                        else if (cell.CellStyle.DataFormat == 9)
                        {
                            newText = (cell.NumericCellValue * 100) + " %";
                        }
                        else if (cell.CellStyle.DataFormat == 0)
                        {
                            newText = cell.NumericCellValue.ToString();
                        }
                        else
                        {
                            // Строки форматирования чисел xlsx в основном совпадают с .net
                            // несовпадения решать по мере появления проблем
                            var formatString = cell.CellStyle.GetDataFormatString();
                            newText = string.IsNullOrWhiteSpace(formatString)
                                ? cell.NumericCellValue.ToString()
                                : cell.NumericCellValue.ToString(formatString);
                        }

                        return newText;
                    }
                default:
                    throw new NotSupportedException();
            }
        }
    }
}
