using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using Cell = MigraDoc.DocumentObjectModel.Tables.Cell;
using HeaderFooter = MigraDoc.DocumentObjectModel.HeaderFooter;
using PageSetup = MigraDoc.DocumentObjectModel.PageSetup;
using Table = MigraDoc.DocumentObjectModel.Tables.Table;

namespace DocxToPdfConverter
{
    public static class TableProcessor
    {
        public static void ProcessTable(XWPFTable table, DocumentObject docObj)
        {
            if (table?.Rows == null || table.Rows.Count == 0 || table.Rows[0] == null)
                return;

            Table mdTable;
            PageSetup pageSetup;

            switch (docObj)
            {
                case Section section:
                {
                    mdTable = section.AddTable();
                    pageSetup = section.PageSetup;
                    break;
                }
                case HeaderFooter footer:
                    mdTable = footer.AddTable();
                    pageSetup = footer.Section!.PageSetup;
                    break;
                default:
                    return;
            }

            SetBorders(table, mdTable);

            mdTable.Rows.Alignment = RowAlignment.Left;
            mdTable.Rows.LeftIndent = Unit.Zero;

            int colCount = CalculateColCount(table);

            var grid = table.GetCTTbl().tblGrid;
            var colWidthsPt = new List<double>();
            const double charToPt = 7.5; // 1 символ ≈ 6pt (0.21см)

            bool hasGrid = grid is { gridCol.Count: > 0 };
            if (hasGrid)
            {
                colWidthsPt.AddRange(grid.gridCol.Select(col => col.w / 20.0));
            }
            else
            {
                // --- Новый алгоритм: ширина по максимальной длине текста ---
                const double minWidthPt = 71;  // 2.5 см
                const double maxWidthPt = 226; // 8 см

                var maxLens = CollectMaxLensByColumns(table, colCount);
                double totalLen = maxLens.Sum();

                // 1. Считаем "естественную" ширину
                for (int c = 0; c < colCount; c++)
                {
                    double w = Math.Max(minWidthPt, Math.Min(maxWidthPt, maxLens[c] * charToPt));
                    colWidthsPt.Add(w);
                }

                // 2. Если сумма меньше ширины страницы — распределяем остаток пропорционально длине текста
                var pageWidthPt = (pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin).Point;
                double totalWidthPt = colWidthsPt.Sum();
                if (totalWidthPt < pageWidthPt && totalLen > 0)
                {
                    double extra = pageWidthPt - totalWidthPt;
                    for (int c = 0; c < colCount; c++)
                    {
                        double add = extra * (maxLens[c] / totalLen);
                        colWidthsPt[c] += add;
                    }
                }

                // 3. Если сумма больше ширины страницы — масштабируем
                totalWidthPt = colWidthsPt.Sum();
                if (totalWidthPt > pageWidthPt)
                {
                    double scale = pageWidthPt / totalWidthPt;
                    for (int i = 0; i < colWidthsPt.Count; i++)
                        colWidthsPt[i] *= scale;
                }
            }

            // Убираем жёсткое масштабирование: если таблица не влезает — пусть выходит за границы
            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromPoint(colWidthsPt[c]));

            //Соберем карту объединения ячеек по вертикали
            var vMerges = CollectVerticalMerges(table, colCount);

            bool isFirstRow = true;

            var r = 0;
            foreach (var row in table.Rows)
            {
                var mdRow = mdTable.AddRow();
                if (isFirstRow)
                {
                    mdRow.HeadingFormat = true;
                    isFirstRow = false;
                }

                mdRow.VerticalAlignment = VerticalAlignment.Center;

                var colNumber = 0;
                foreach (var cell in row.GetTableCells())
                {
                    var cellProps = cell.GetCTTc()?.tcPr;

                    var mdCell = mdRow.Cells[colNumber];
                    var vMerge = vMerges[r, colNumber];
                    var colWidth = colWidthsPt[colNumber];

                    var hasGridSpan = int.TryParse(cellProps?.gridSpan?.val, out var span);
                    colNumber += hasGridSpan ? span : 1;

                    if (hasGridSpan)
                    {
                        mdCell.MergeRight = span - 1;
                    }

                    mdCell.MergeDown = vMerge;

                    mdCell.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

                    foreach (var elem in cell.BodyElements)
                    {
                        switch (elem)
                        {
                            case XWPFParagraph para:
                                var maxWordLen = (int)(colWidth / charToPt);
                                    ParagraphProcessor.ProcessParagraph(para, mdCell, maxWordLen);
                                 break;
                            case XWPFTable nestedTable:
                                ProcessTable(nestedTable, mdCell);
                                break;
                        }
                    }
                }

                r++;
            }
        }

        private static int CalculateColCount(XWPFTable table)
        {
            return table
                .Rows[0]
                .GetTableCells()
                .Sum(cell => int.TryParse(cell?.GetCTTc()?.tcPr?.gridSpan?.val, out var hMerge) ? hMerge : 1);
        }

        private static int[,] CollectVerticalMerges(XWPFTable table, int colCount)
        {
            var vMerges = new int[table.Rows.Count, colCount];
            var r = 0;
            foreach (var row in table.Rows)
            {
                var c = 0;
                foreach (var cell in row.GetTableCells())
                {
                    var cellProps = cell?.GetCTTc()?.tcPr;

                    switch (cellProps?.vMerge?.val)
                    {
                        case ST_Merge.restart:
                            vMerges[r, c] = 1;
                            break;
                        case ST_Merge.@continue:
                            int pr;
                            for (pr = r-1; pr > 0 && vMerges[pr, c] == 0 ; pr--) { }
                            vMerges[pr, c]++;
                            break;
                    }

                    c += int.TryParse(cellProps?.gridSpan?.val, out var hMerge) ? hMerge : 1;
                }

                r++;
            }

            return vMerges;
        }

        private static List<int> CollectMaxLensByColumns(XWPFTable table, int colCount)
        {
            var maxLens = new List<int>(new int[colCount]);

            foreach (var row in table.Rows)
            {
                for (int c = 0; c < colCount; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell != null)
                    {
                        int len = 0;
                        foreach (var elem in cell.BodyElements)
                        {
                            if (elem is XWPFParagraph para)
                                len += para.ParagraphText?.Length ?? 0;
                        }

                        maxLens[c] = Math.Max(maxLens[c], len);
                    }
                }
            }

            return maxLens;
        }

        private static void SetBorders(XWPFTable table, Table mdTable)
        {
            var tblBorders = table.GetCTTbl().tblPr.tblBorders;

            mdTable.Borders.Visible = true;
            if (tblBorders is not null)
            {
                mdTable.Borders.Visible = tblBorders.insideH.sz > 0 || tblBorders.insideV.sz > 0
                    || tblBorders.bottom.sz > 0
                    || tblBorders.top.sz > 0 || tblBorders.left.sz > 0 || tblBorders.right.sz > 0;
            }
        }

        private static void ProcessTable(XWPFTable table, Cell cell)
        {
            if (table?.Rows == null || table.Rows.Count == 0 || table.Rows[0] == null)
                return;

            var mdTable = cell.Elements.AddTable();
            mdTable.Rows.Alignment = RowAlignment.Left;
            mdTable.Rows.LeftIndent = Unit.Zero;

            int colCount = table.Rows[0].GetTableCells().Count;
            var grid = table.GetCTTbl().tblGrid;
            var colWidthsPt = new List<double>();
            bool hasGrid = grid is { gridCol.Count: > 0 };

            if (hasGrid)
            {
                colWidthsPt.AddRange(grid.gridCol.Select(col => col.w / 20.0));

                while (colWidthsPt.Count < colCount)
                {
                    colWidthsPt.Add(42); // 1.5 см = 42 pt
                }
            }
            else
            {
                const double minWidthPt = 42;
                const double maxWidthPt = 226;
                const double charToPt = 6.0;

                var maxLens = new List<int>(new int[colCount]);
                foreach (var row in table.Rows)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        var cell2 = row.GetCell(c);
                        if (cell2 != null)
                        {
                            int len = 0;
                            foreach (var elem in cell2.BodyElements)
                            {
                                if (elem is XWPFParagraph para)
                                    len += para.ParagraphText?.Length ?? 0;
                            }

                            maxLens[c] = Math.Max(maxLens[c], len);
                        }
                    }
                }

                for (int c = 0; c < colCount; c++)
                {
                    double w = Math.Max(minWidthPt, Math.Min(maxWidthPt, maxLens[c] * charToPt));
                    colWidthsPt.Add(w);
                }
            }

            for (int i = 0; i < colWidthsPt.Count; i++)
            {
                colWidthsPt[i] = Math.Max(42, Math.Min(colWidthsPt[i], 226));
            }

            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromPoint(colWidthsPt[c]));

            foreach (var row in table.Rows)
            {
                var mdRow = mdTable.AddRow();
                for (int c = 0; c < colCount; c++)
                {
                    var cell2 = row.GetCell(c);
                    var mdCell = mdRow.Cells[c];
                    if (cell2 != null)
                    {
                        // Включаем перенос текста в ячейке
                        mdCell.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                        foreach (var elem in cell2.BodyElements)
                        {
                            switch (elem)
                            {
                                case XWPFParagraph para:
                                    ParagraphProcessor.ProcessParagraph(para, mdCell, 0);
                                    break;
                                case XWPFTable nestedTable:
                                    ProcessTable(nestedTable, mdCell);
                                    break;
                            }
                        }
                    }
                }
            }
        }
    }
}