using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.XWPF.UserModel;

namespace DocxToPdfConverter
{
    public static class TableProcessor
    {
        public static void ProcessTable(XWPFTable table, Section section)
        {
            var mdTable = section.AddTable();
            mdTable.Borders.Visible = true;
            mdTable.Rows.Alignment = RowAlignment.Center;
            int colCount = table.Rows[0].GetTableCells().Count;
            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromCentimeter(4));
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
            var mdTable = cell.Elements.AddTable();
            int colCount = table.Rows[0].GetTableCells().Count;
            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromCentimeter(4));
            foreach (var row in table.Rows)
            {
                var mdRow = mdTable.AddRow();
                for (int c = 0; c < colCount; c++)
                {
                    var cell2 = row.GetCell(c);
                    var mdCell = mdRow.Cells[c];
                    if (cell2 != null)
                    {
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