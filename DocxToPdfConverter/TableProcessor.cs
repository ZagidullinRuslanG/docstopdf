using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.XWPF.UserModel;

namespace DocxToPdfConverter
{
    public static class TableProcessor
    {
        public static void ProcessTable(XWPFTable table, Section section)
        {
            if (table == null || table.Rows == null || table.Rows.Count == 0 || table.Rows[0] == null)
                return;
            var mdTable = section.AddTable();
            mdTable.Borders.Visible = true;
            mdTable.Rows.Alignment = RowAlignment.Left;
            mdTable.Rows.LeftIndent = Unit.Zero;
            int colCount = table.Rows[0].GetTableCells().Count;
            var grid = table.GetCTTbl().tblGrid;
            List<double> colWidthsPt = new List<double>();
            bool hasGrid = grid != null && grid.gridCol != null && grid.gridCol.Count > 0;
            if (hasGrid)
            {
                foreach (var col in grid.gridCol)
                {
                    double widthPt = col.w / 20.0; // twips to points
                    colWidthsPt.Add(widthPt);
                }
                while (colWidthsPt.Count < colCount)
                {
                    colWidthsPt.Add(71); // 2.5 см = 71 pt
                }
            }
            else
            {
                // --- Новый алгоритм: ширина по максимальной длине текста ---
                double minWidthPt = 71; // 2.5 см
                double maxWidthPt = 226; // 8 см
                double charToPt = 6.0; // 1 символ ≈ 6pt (0.21см)
                int rowCount = table.Rows.Count;
                List<int> maxLens = new List<int>(new int[colCount]);
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
                double totalLen = maxLens.Sum();
                // 1. Считаем "естественную" ширину
                for (int c = 0; c < colCount; c++)
                {
                    double w = Math.Max(minWidthPt, Math.Min(maxWidthPt, maxLens[c] * charToPt));
                    colWidthsPt.Add(w);
                }
                // 2. Если сумма меньше ширины страницы — распределяем остаток пропорционально длине текста
                var pageWidthPt = (section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin).Point;
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
            // Ограничения min/max ширины (1.5см=42pt, 8см=226pt)
            for (int i = 0; i < colWidthsPt.Count; i++)
            {
                colWidthsPt[i] = Math.Max(42, Math.Min(colWidthsPt[i], 226));
            }
            // Убираем жёсткое масштабирование: если таблица не влезает — пусть выходит за границы
            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromPoint(colWidthsPt[c]));
            bool isFirstRow = true;
            foreach (var row in table.Rows)
            {
                var mdRow = mdTable.AddRow();
                if (isFirstRow)
                {
                    mdRow.HeadingFormat = true;
                    isFirstRow = false;
                }
                mdRow.VerticalAlignment = VerticalAlignment.Center;
                for (int c = 0; c < colCount; c++)
                {
                    var cell = row.GetCell(c);
                    var mdCell = mdRow.Cells[c];
                    if (cell != null)
                    {
                        // Включаем перенос текста в ячейке
                        mdCell.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                        var shd = cell.GetCTTc().tcPr?.shd;
                        if (shd != null && !string.IsNullOrEmpty(shd.fill) && shd.fill != "auto" && shd.fill.Length == 6)
                        {
                            try
                            {
                                int rCol = System.Convert.ToInt32(shd.fill.Substring(0, 2), 16);
                                int gCol = System.Convert.ToInt32(shd.fill.Substring(2, 2), 16);
                                int bCol = System.Convert.ToInt32(shd.fill.Substring(4, 2), 16);
                                mdCell.Shading.Color = Color.FromRgb((byte)rCol, (byte)gCol, (byte)bCol);
                            }
                            catch { }
                        }
                        foreach (var elem in cell.BodyElements)
                        {
                            if (elem is XWPFParagraph para)
                                ParagraphProcessor.ProcessParagraph(para, mdCell);
                            else if (elem is XWPFTable nestedTable)
                                ProcessTable(nestedTable, mdCell);
                        }
                    }
                }
            }
        }

        public static void ProcessTable(XWPFTable table, Cell cell)
        {
            if (table == null || table.Rows == null || table.Rows.Count == 0 || table.Rows[0] == null)
                return;
            var mdTable = cell.Elements.AddTable();
            mdTable.Rows.Alignment = RowAlignment.Left;
            mdTable.Rows.LeftIndent = Unit.Zero;
            int colCount = table.Rows[0].GetTableCells().Count;
            var grid = table.GetCTTbl().tblGrid;
            List<double> colWidthsPt = new List<double>();
            bool hasGrid = grid != null && grid.gridCol != null && grid.gridCol.Count > 0;
            if (hasGrid)
            {
                foreach (var col in grid.gridCol)
                {
                    double widthPt = col.w / 20.0;
                    colWidthsPt.Add(widthPt);
                }
                while (colWidthsPt.Count < colCount)
                {
                    colWidthsPt.Add(42); // 1.5 см = 42 pt
                }
            }
            else
            {
                double minWidthPt = 42;
                double maxWidthPt = 226;
                double charToPt = 6.0;
                int rowCount = table.Rows.Count;
                List<int> maxLens = new List<int>(new int[colCount]);
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
                            if (elem is XWPFParagraph para)
                                ParagraphProcessor.ProcessParagraph(para, mdCell);
                            else if (elem is XWPFTable nestedTable)
                                ProcessTable(nestedTable, mdCell);
                        }
                    }
                }
            }
        }
    }
} 