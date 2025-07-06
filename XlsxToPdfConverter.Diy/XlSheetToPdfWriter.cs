using System;
using System.Collections.Generic;
using System.Linq;
using NpoiHelpers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Преобразование страницы Excel в PDF.
    /// </summary>
    internal class XlSheetToPdfWriter
    {
        private readonly PdfWriter pdfWriter;
        private readonly IWorkbook workbook;
        private readonly ISheet sheet;
        private XlSheetToPdfWriterProperties properties;
        private readonly int sheetId;
        private readonly MergedCellsData mergedCellsData;

        private readonly int colNum;

        private readonly float[] colPos;
        private readonly float[] colWidth;
        private readonly PdfPartitionLayout layout;

        private XlSheetToPdfWriterProperties GetSheetProperties(ISheet sheet)
        {
            short lastColNum = 0;
            for (int ri = 0; ri < sheet.LastRowNum; ri++)
            {
                IRow row = sheet.GetRow(ri);
                if (row != null && row.LastCellNum > lastColNum)
                {
                    lastColNum = row.LastCellNum;
                }
            }

            XlSheetToPdfWriterProperties sheetProperties =
                new XlSheetToPdfWriterProperties()
                {
                    ColNum = lastColNum,
                    PagePrintSize = PrintDataConverter.GetPaperSize(sheet.PrintSetup),
                    PaperOrientation = PrintDataConverter.GetOrientation(sheet.PrintSetup),
                    PagePrintMargins = PrintDataConverter.GetMargin(sheet),
                    PageHeaderAndFooterMargins = PrintDataConverter.GetHeaderAndFooterMargin(sheet),
                };

            return sheetProperties;
        }

        public XlSheetToPdfWriter(
            IWorkbook workbook,
            PdfWriter pdfWriter,
            ISheet sheet,
            int sheetId,
            MergedCellsData mergedCellsData)
        {
            this.workbook = workbook;
            this.sheet = sheet;
            this.sheetId = sheetId;
            this.mergedCellsData = mergedCellsData;
            this.pdfWriter = pdfWriter;
            this.properties = GetSheetProperties(sheet);

            this.colNum = properties.ColNum;

            // В некоторых случаях NPOI не правильно определяет ширину колонок по умолчанию
            if (sheet.DefaultColumnWidth == 0)
            {
                sheet.DefaultColumnWidth = 8;
            }

            (colPos, colWidth) = GetColumnsSize(sheet, colNum);

            double xlActualWidth = colPos[colNum - 1] + colWidth[colNum - 1];
            layout = new PdfPartitionLayout(
                xlActualWidth,
                properties.PagePrintSize,
                properties.PaperOrientation,
                properties.PagePrintMargins,
                properties.PageHeaderAndFooterMargins);
        }

        private void WriteHeaderAndFooter(ISheet sheet, int startPage, int endPage)
        {
            // todo: потенциально возможны листы с нечетными колонтитулами или на первой странице
            IHeader head = sheet.Header;
            IFooter foot = sheet.Footer;
            for (int i = startPage; i <= endPage; i++)
            {
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.TopLeft, head.Left);
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.TopRight, head.Right);
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.TopCenter, head.Center);
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.BottomLeft, foot.Left);
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.BottomRight, foot.Right);
                pdfWriter.WriteHeaderOrFooter(i, HeaderOrFooterPosition.BottomCenter, foot.Center);
            }
        }

        public void Write()
        {
            pdfWriter.AddPartition(layout);

            var processedMerges = new HashSet<int>();
            int startPage = pdfWriter.GetCurrentPageNum();
            double rowPos = 0;

            WriteBackground();

            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                int nextRangeFrom = 0;
                ICell waitingCell = null;

                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    rowPos += sheet.DefaultRowHeightInPoints;
                    continue;
                }

                if (row.Hidden.HasValue && row.Hidden.Value)
                {
                    continue;
                }

                // Сначала проходим все ячейки, чтобы найти самую высокую объединенную ячейку и по ней
                // задать возможный перенос всего блока на следующую страницу PDF.
                var mergesAreas = new Dictionary<int, Rect>();
                int maxHeightMergeIndex = -1;
                for (int columnIndex = 0; columnIndex < colNum; columnIndex++)
                {
                    ICell cell = row.GetCell(columnIndex);
                    // note: Специально не используем ICell.IsMergedCell, т.к. работает чрезвычайно медленно.
                    if (cell == null || !mergedCellsData.IsMergedCell(sheetId, rowIndex, columnIndex))
                    {
                        continue;
                    }

                    (CellRangeAddress range, int mergeIndex) = mergedCellsData.GetMergedRegion(sheetId, rowIndex, columnIndex);
                    if (processedMerges.Contains(mergeIndex))
                    {
                        continue;
                    }

                    double mergeWidth = colWidth
                        .Where((val, ind) => ind >= range.FirstColumn && ind <= range.LastColumn).Sum();
                    double mergeHeight = SumRowHeights(sheet, range.FirstRow, range.LastRow);
                    mergesAreas.Add(
                        mergeIndex,
                        new Rect(
                            colPos[range.FirstColumn],
                            rowPos,
                            mergeWidth,
                            mergeHeight));
                    if (maxHeightMergeIndex < 0 || mergesAreas[maxHeightMergeIndex].H < mergeHeight)
                    {
                        maxHeightMergeIndex = mergeIndex;
                    }

                    columnIndex += range.LastColumn - range.FirstColumn;
                }

                // Теперь проходим ячейки повторно, уже записывая все данные.
                for (int columnIndex = 0; columnIndex < colNum; columnIndex++)
                {
                    ICell cell = row.GetCell(columnIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    ICellStyle style = cell.CellStyle;
                    StyleList styleList = new StyleList(cell.CellStyle.GetFont(workbook));
                    if (cell.CellType is CellType.String or CellType.Formula or CellType.Blank)
                    {
                        var richText = (XSSFRichTextString)cell.RichStringCellValue;

                        int formattingRuns = cell.RichStringCellValue.NumFormattingRuns;
                        for (int i = 1; i < formattingRuns; i++)
                        {
                            styleList.AddStyle(richText.GetIndexOfFormattingRun(i),
                                richText.GetLengthOfFormattingRun(i),
                                richText.GetFontOfFormattingRun(i));
                        }
                    }

                    // note: Специально не используем ICell.IsMergedCell, т.к. работает чрезвычайно медленно.
                    if (mergedCellsData.IsMergedCell(sheetId, rowIndex, columnIndex))
                    {
                        (CellRangeAddress range, int mergeIndex) =
                            mergedCellsData.GetMergedRegion(sheetId, rowIndex, columnIndex);

                        // Записываем ожидающие данные.
                        Write(rowPos, row, nextRangeFrom, columnIndex - 1, waitingCell);
                        waitingCell = null;

                        nextRangeFrom = range.LastColumn + 1;

                        if (processedMerges.Contains(mergeIndex))
                        {
                            pdfWriter.WriteBorders(style,
                                new Rect(colPos[columnIndex], rowPos, colWidth[columnIndex], row.HeightInPoints), true);
                            continue;
                        }

                        processedMerges.Add(mergeIndex);

                        if (!cell.IsEmpty())
                        {
                            // Текст в объединенной области всегда в рамках границы, поэтому записываем сразу.
                            if (style.Alignment == HorizontalAlignment.General &&
                                cell.IsNumber())
                            {
                                style.Alignment = HorizontalAlignment.Right;
                            }

                            Rect mArea = mergesAreas[mergeIndex];
                            if (mArea.W != 0)
                            {
                                pdfWriter.WriteText(
                                    cell.AsText(),
                                    style,
                                    styleList,
                                    mArea);
                            }
                        }

                        pdfWriter.WriteBorders(style, mergesAreas[mergeIndex]);

                        // Иногда данные по границам объединенных ячеек хранятся у последних ячеек
                        int lastCellId = columnIndex + (range.LastColumn - range.FirstColumn);
                        ICell lastCell = row.GetCell(lastCellId);
                        if (lastCell != null && lastCellId != columnIndex)
                        {
                            pdfWriter.WriteBorders(
                                lastCell.CellStyle,
                                new Rect(colPos[lastCellId], rowPos, colWidth[lastCellId], row.HeightInPoints),
                                true);
                        }

                        columnIndex += range.LastColumn - range.FirstColumn;
                        continue;
                    }

                    if (cell.IsEmpty())
                    {
                        pdfWriter.WriteBorders(
                            style, new Rect(colPos[columnIndex], rowPos, colWidth[columnIndex], row.HeightInPoints));
                        continue;
                    }

                    pdfWriter.WriteBorders(
                        style, new Rect(colPos[columnIndex], rowPos, colWidth[columnIndex], row.HeightInPoints));

                    // Записываем ожидающие данные.
                    if (Write(rowPos, row, nextRangeFrom, columnIndex - 1, waitingCell))
                    {
                        waitingCell = null;
                        nextRangeFrom = columnIndex;
                    }

                    if (style.WrapText)
                    {
                        // Записываем текущую ячейку.
                        Write(rowPos, row, columnIndex, columnIndex, cell);

                        nextRangeFrom = columnIndex + 1;
                    }
                    else
                    {
                        switch (style.Alignment)
                        {
                            case HorizontalAlignment.Center:
                                waitingCell = cell;
                                break;
                            case HorizontalAlignment.Left:
                                waitingCell = cell;
                                nextRangeFrom = columnIndex;
                                break;
                            case HorizontalAlignment.General:
                                if (cell.IsNumber())
                                {
                                    style.Alignment = HorizontalAlignment.Right;
                                    // Записываем текущий текст.
                                    Write(rowPos, row, columnIndex, columnIndex, cell);
                                    nextRangeFrom = columnIndex + 1;
                                }
                                else
                                {
                                    waitingCell = cell;
                                    nextRangeFrom = columnIndex;
                                }

                                break;
                            case HorizontalAlignment.Right:
                                // Записываем текущий текст.
                                Write(rowPos, row, columnIndex, columnIndex, cell);
                                nextRangeFrom = columnIndex + 1;
                                break;
                            default:
                                throw new NotSupportedException();
                        }
                    }
                }

                Write(rowPos, row, nextRangeFrom, colNum - 1, waitingCell);

                rowPos += row.HeightInPoints;
            }

            WriteHeaderAndFooter(sheet, startPage, pdfWriter.GetCurrentPageNum());
        }

        private void WriteBackground()
        {
            var processedMerges = new HashSet<int>();
            double rowPos = 0;
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    rowPos += sheet.DefaultRowHeightInPoints;
                    continue;
                }
                
                if (row.Hidden.HasValue && row.Hidden.Value)
                {
                    continue;
                }

                // Сначала проходим все ячейки, чтобы найти самую высокую объединенную ячейку и по ней
                // задать возможный перенос всего блока на следующую страницу PDF.
                var mergesAreas = new Dictionary<int, Rect>();
                int maxHeightMergeIndex = -1;
                for (int columnIndex = 0; columnIndex < colNum; columnIndex++)
                {
                    ICell cell = row.GetCell(columnIndex);
                    // note: Специально не используем ICell.IsMergedCell, т.к. работает чрезвычайно медленно.
                    if (cell == null || !mergedCellsData.IsMergedCell(sheetId, rowIndex, columnIndex))
                    {
                        continue;
                    }

                    (CellRangeAddress range, int mergeIndex) = mergedCellsData.GetMergedRegion(sheetId, rowIndex, columnIndex);
                    if (processedMerges.Contains(mergeIndex))
                    {
                        continue;
                    }

                    double mergeWidth = colWidth
                        .Where((val, ind) => ind >= range.FirstColumn && ind <= range.LastColumn).Sum();
                    double mergeHeight = SumRowHeights(sheet, range.FirstRow, range.LastRow);
                    mergesAreas.Add(
                        mergeIndex,
                        new Rect(
                            colPos[range.FirstColumn],
                            rowPos,
                            mergeWidth,
                            mergeHeight));
                    if (maxHeightMergeIndex < 0 || mergesAreas[maxHeightMergeIndex].H < mergeHeight)
                    {
                        maxHeightMergeIndex = mergeIndex;
                    }

                    columnIndex += range.LastColumn - range.FirstColumn;
                }

                if (maxHeightMergeIndex >= 0)
                {
                    pdfWriter.SetNotSplittableByPagesArea(mergesAreas[maxHeightMergeIndex]);
                }

                // Теперь проходим ячейки повторно, уже записывая все данные.
                for (int columnIndex = 0; columnIndex < colNum; columnIndex++)
                {
                    ICell cell = row.GetCell(columnIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    ICellStyle style = cell.CellStyle;

                    // note: Специально не используем ICell.IsMergedCell, т.к. работает чрезвычайно медленно.
                    if (mergedCellsData.IsMergedCell(sheetId, rowIndex, columnIndex))
                    {
                        (CellRangeAddress range, int mergeIndex) =
                            mergedCellsData.GetMergedRegion(sheetId, rowIndex, columnIndex);

                        if (processedMerges.Contains(mergeIndex))
                        {
                            continue;
                        }

                        processedMerges.Add(mergeIndex);

                        pdfWriter.WriteBackground(style, mergesAreas[mergeIndex]);

                        columnIndex += range.LastColumn - range.FirstColumn;
                        continue;
                    }

                    pdfWriter.WriteBackground(
                        style, new Rect(colPos[columnIndex], rowPos, colWidth[columnIndex], row.HeightInPoints));
                }

                rowPos += row.HeightInPoints;
            }
        }

        private bool Write(
            double rowPos,
            IRow row,
            int colFrom,
            int colTo,
            ICell cell)
        {
            if (cell == null)
            {
                return false;
            }

            if (cell.IsEmpty())
            {
                return false;
            }

            // note: используется для записи текста, выровненного по центру и выходящего за границы ячейки.
            // В Excel в этом случае выравнивание идет по центры ячейки с данными, но текст обрезается по ближайшей левой и правой с данными.
            // Тут реализована обрезка по минимальному расстоянию до левой или правой с обоих сторон.
            if (cell.CellStyle.Alignment == HorizontalAlignment.Center)
            {
                // Делаем число ячеек слева и справа от ячейки с данными одинаковым.

                int left = cell.ColumnIndex - colFrom;
                int right = colTo - cell.ColumnIndex;
                if (left != right)
                {
                    if (left < right)
                    {
                        colTo -= (right - left);
                    }
                    else
                    {
                        colFrom += (left - right);
                    }
                }
            }

            var width = colPos[colTo] + colWidth[colTo] - colPos[colFrom];
            if (width == 0)
            {
                return false;
            }

            ICellStyle style = cell.CellStyle;
            IFont font = style.GetFont(workbook);
            pdfWriter.WriteText(
                cell.AsText(),
                style,
                font,
                new Rect(colPos[colFrom], rowPos, width, row.HeightInPoints));

            return true;
        }

        private static (float[] pos, float[] width) GetColumnsSize(ISheet sheet, int colNum)
        {
            var pos = new float[colNum];
            var width = new float[colNum];
            pos[0] = 0;
            width[0] = (float)(sheet.IsColumnHidden(0) ? 0 : sheet.GetColumnWidthInPixels(0));
            for (int ci = 1; ci < colNum; ci++)
            {
                pos[ci] = pos[ci - 1] + width[ci - 1];
                width[ci] = (float)(sheet.IsColumnHidden(ci) ? 0 : sheet.GetColumnWidthInPixels(ci));
            }

            return (pos, width);
        }

        private double SumRowHeights(ISheet sheet, int firstRow, int lastRow)
        {
            double sum = 0;
            for (int i = firstRow; i <= lastRow; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    sum += row.HeightInPoints;
                }
            }

            return sum;
        }
    }
}