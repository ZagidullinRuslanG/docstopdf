using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace NpoiHelpers
{
    public static class SheetExtensions
    {
        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, double? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, double? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnIndex, value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnIndex, "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, string address, float value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, double value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, int? value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, int value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, decimal value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, decimal? value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, DateTime? value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, string value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValueAndHeight(this ISheet sheet, string address, string value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValueAndHeight(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, float value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, double value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, double value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, float value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, string value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValueAndHeight(this ISheet sheet, int rowIndex, int columnIndex, string value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            cell.CellStyle.WrapText = true;
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, string value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValueAndHeight(this ISheet sheet, int rowIndex, string columnName, string value)
        {
            return sheet.SetCellValueAndHeight(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, int value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, decimal value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(Convert.ToDouble(value));
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, decimal value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValueAndHeight(this ISheet sheet, int rowIndex, string columnName, IRichTextString value)
        {
            return sheet.SetCellValueAndHeight(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, IRichTextString value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, decimal? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, decimal? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnIndex, value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnIndex, "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, DateTime value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, DateTime? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnIndex, value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnIndex, "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, int value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValueAndHeight(this ISheet sheet, int rowIndex, int columnIndex, IRichTextString value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            var oldHeight = sheet.GetRow(rowIndex).Height;
            var newHeight = (short)( Html2RtfConverter.GetOptimalHeight(cell) * 20);
            sheet.GetRow(rowIndex).Height = oldHeight > newHeight ? oldHeight : newHeight;
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, IRichTextString value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, int? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnIndex, value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnIndex, "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, DateTime value)
        {
            return sheet.SetCellValue(rowIndex, columnName.XlCol(), value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, DateTime? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), "-");
            }
        }

        public static ICell SetCellStyle(this ISheet sheet, int rowIndex, string col, ICellStyle style)
        {
            return sheet.SetCellStyle(rowIndex, col.XlCol(), style);
        }
        
        public static void SetCellRangeStyle(this ISheet sheet, int firstRowNum, int lastRowNum, int firstColNum, int lastColNum,  ICellStyle cellStyle)
        {
            for (var i = firstRowNum; i <= lastRowNum; i++)
            {
                for (var j = firstColNum; j <= lastColNum; j++)
                { 
                    sheet.SetCellStyle(i, j, cellStyle);
                }
            }
        }
        public static ICell SetCellStyle(this ISheet sheet, int rowIndex, int columnIndex, ICellStyle style)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.CellStyle = style;
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, string address, long? value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, string address, long value)
        {
            CellRangeAddress range = CellRangeAddress.ValueOf(address);
            return sheet.SetCellValue(range.FirstRow, range.FirstColumn, value);
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, long value)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(columnIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, columnIndex);
            }

            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, long? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnIndex, value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnIndex, "");
            }
        }

        public static ICell SetCellValue(this ISheet sheet, int rowIndex, string columnName, long? value)
        {
            if (value.HasValue)
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), value.Value);
            }
            else
            {
                return sheet.SetCellValue(rowIndex, columnName.XlCol(), "");
            }
        }

        public static ICell CreateVerticalMergedCell(this ISheet sheet, string colName, int rowIndex1, int rowIndex2)
        {
            return sheet.CreateVerticalMergedCell(colName.XlCol(), rowIndex1, rowIndex2);
        }

        public static ICell CreateVerticalMergedCell(this ISheet sheet, int colIndex, int rowIndex1, int rowIndex2)
        {
            ICell cell = sheet.GetRow(rowIndex1).GetCell(colIndex);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex1, colIndex);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex1, rowIndex2, colIndex, colIndex));

            return cell;
        }

        public static ICell CreateHorizontalMergedCell(this ISheet sheet, int rowIndex, int colIndex1, int colIndex2)
        {
            ICell cell = sheet.GetRow(rowIndex).GetCell(colIndex1);
            if (cell == null)
            {
                throw new ExcelCellNotFoundException(rowIndex, colIndex1);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, colIndex1, colIndex2));

            return cell;
        }
    }

    // Заглушка для Html2RtfConverter
    public static class Html2RtfConverter
    {
        public static string Convert(string html) => string.Empty;

        public static int GetOptimalHeight(string html) => 0;
        public static int GetOptimalHeight(NPOI.SS.UserModel.ICell cell) => 0;
    }
}
