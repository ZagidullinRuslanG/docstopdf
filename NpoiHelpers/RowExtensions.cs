using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace NpoiHelpers
{
    public static class RowExtensions
    {
        public static ICell CreateCell(this IRow row, int column, ICellStyle style, CellType type)
        {
            ICell cell = row.CreateCell(column);
            cell.CellStyle = style;
            cell.SetCellType(type);
            return cell;
        }

        public static ICell CreateCell(this IRow row, int column, ICellStyle style, string value)
        {
            ICell cell = row.CreateCell(column, style, CellType.String);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell CreateCell(this IRow row, string column, ICellStyle style, string value)
        {
            return row.CreateCell(column.XlCol(), style, value);
        }

        public static ICell CreateCell(this IRow row, int column, ICellStyle style, DateTime value)
        {
            ICell cell = row.CreateCell(column, style, CellType.String);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell CreateCell(this IRow row, string column, ICellStyle style, DateTime value)
        {
            return row.CreateCell(column.XlCol(), style, value);
        }


        public static ICell CreateCell(this IRow row, int column, ICellStyle style, double value)
        {
            ICell cell = row.CreateCell(column, style, CellType.Numeric);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell CreateMergedCell(this IRow row, int col1, int col2, ICellStyle style, string value)
        {
            ICell cell = row.CreateCell(col1, style, value);
            for (int colNum = col1 + 1; colNum <= col2; colNum++)
            {
                row.CreateCell(colNum, style, " ");
            }

            row.Sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, col1, col2));
            return cell;
        }

        public static ICell CreateMergedCell(this IRow row, int col1, int col2, ICellStyle style, DateTime value)
        {
            ICell cell = row.CreateCell(col1, style, value);
            row.Sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, col1, col2));
            return cell;
        }

        public static ICell CreateMergedCell(this IRow row, string col1, string col2, ICellStyle style, string value)
        {
            return row.CreateMergedCell(col1.XlCol(), col2.XlCol(), style, value);
        }

        public static ICell CreateMergedCell(this IRow row, string col1, string col2, ICellStyle style, DateTime value)
        {
            return row.CreateMergedCell(col1.XlCol(), col2.XlCol(), style, value);
        }
    }
}
