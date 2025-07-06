using System;

namespace NpoiHelpers
{
    public class ExcelCellNotFoundException : Exception
    {
        public int RowIndex { get; private set; }

        public int ColumnIndex { get; private set; }

        public ExcelCellNotFoundException(int rowIndex, int columnIndex) : base(
                $"Не найдена ячейка в Excel {columnIndex.XlCol()}{rowIndex + 1} " +
                $"[RowIndex = {rowIndex}, ColumnIndex = {columnIndex}]")
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
    }
}