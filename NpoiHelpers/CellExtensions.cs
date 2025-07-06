using NPOI.SS.UserModel;

namespace NpoiHelpers
{
    public static class CellExtensions
    {
        public static ICell SetStyle(this ICell cell, ICellStyle style)
        {
            cell.CellStyle = style;
            return cell;
        }
    }
}