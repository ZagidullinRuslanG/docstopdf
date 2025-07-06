namespace XlsxToPdfConverter.Diy
{
    using NPOI.SS.Util;

    public class MergedCellsData
    {
        // Заглушка для совместимости. Реализуйте по необходимости.

        public bool IsMergedCell(int sheetId, int rowIndex, int columnIndex)
        {
            // Заглушка: всегда возвращает false
            return false;
        }

        public (CellRangeAddress, int) GetMergedRegion(int sheetId, int rowIndex, int columnIndex)
        {
            // Заглушка: возвращает пустой CellRangeAddress и -1
            return (new CellRangeAddress(0, 0, 0, 0), -1);
        }
    }
} 