using PdfSharp.Drawing;
using NPOI.OpenXmlFormats.Wordprocessing;
using PdfSharp.Fonts;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using NPOI.XWPF.UserModel;
using Document = MigraDoc.DocumentObjectModel.Document;
using ParagraphAlignment = NPOI.XWPF.UserModel.ParagraphAlignment;
using System.Reflection;
using MigraDoc.DocumentObjectModel.Tables;

namespace DocxToPdfConverter
{
    public class DocxToPdfConverter
    {
        static DocxToPdfConverter()
        {
            GlobalFontSettings.FontResolver = new MigraDocFontResolver();
        }

        public void Convert(string docxPath, string pdfPath)
        {
            var document = new Document();
            var section = document.AddSection();

            using (var stream = File.OpenRead(docxPath))
            {
                var doc = new XWPFDocument(stream);
                foreach (var element in doc.BodyElements)
                {
                    if (element is XWPFParagraph para)
                    {
                        ProcessParagraph(para, section);
                    }
                    else if (element is XWPFTable table)
                    {
                        ProcessTable(table, section);
                    }
                }
            }
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.Save(pdfPath);
        }

        private void ProcessParagraph(XWPFParagraph para, Section section)
        {
            var mdPara = section.AddParagraph();
            // Выравнивание
            switch (para.Alignment)
            {
                case ParagraphAlignment.CENTER:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center;
                    break;
                case ParagraphAlignment.RIGHT:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Right;
                    break;
                case ParagraphAlignment.BOTH:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify;
                    break;
                default:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                    break;
            }
            // Отступы
            int ToIntSafe(object value)
            {
                if (value == null) return 0;
                if (value is int i) return i;
                if (value is long l) return (int)l;
                if (value is ulong ul) return (int)ul;
                if (value is string s && int.TryParse(s, out int result)) return result;
                return 0;
            }
            int GetSpacingAfterSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.spacing?.after);
            int GetSpacingBeforeSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.spacing?.before);
            int GetIndentationLeftSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.ind?.left);
            int GetIndentationRightSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.ind?.right);
            int indentationLeft = GetIndentationLeftSafe(para);
            int indentationRight = GetIndentationRightSafe(para);
            int spacingBefore = GetSpacingBeforeSafe(para);
            int spacingAfter = GetSpacingAfterSafe(para);
            if (indentationLeft > 0)
                mdPara.Format.LeftIndent = indentationLeft / 20.0;
            if (indentationRight > 0)
                mdPara.Format.RightIndent = indentationRight / 20.0;
            if (spacingBefore > 0)
                mdPara.Format.SpaceBefore = spacingBefore / 20.0;
            if (spacingAfter > 0)
                mdPara.Format.SpaceAfter = spacingAfter / 20.0;
            // Стилизация run'ов
            foreach (var run in para.Runs)
            {
                var mdText = mdPara.AddFormattedText(run.Text);
                if (run.IsBold) mdText.Bold = true;
                if (run.IsItalic) mdText.Italic = true;
                if (run.Underline != UnderlinePatterns.None) mdText.Underline = Underline.Single;
                if (run.FontSize > 0) mdText.Size = run.FontSize;
                if (!string.IsNullOrEmpty(run.FontFamily)) mdText.Font.Name = run.FontFamily;
                // Цвет
                var clr = run.GetCTR().rPr?.color?.val;
                if (!string.IsNullOrEmpty(clr) && clr.Length == 6)
                {
                    try
                    {
                        int rr = System.Convert.ToInt32(clr.Substring(0, 2), 16);
                        int gg = System.Convert.ToInt32(clr.Substring(2, 2), 16);
                        int bb = System.Convert.ToInt32(clr.Substring(4, 2), 16);
                        mdText.Color = MigraDoc.DocumentObjectModel.Color.FromRgb((byte)rr, (byte)gg, (byte)bb);
                    }
                    catch { }
                }
            }
        }

        private void ProcessTable(XWPFTable table, Section section)
        {
            var mdTable = section.AddTable();
            mdTable.Borders.Visible = true; // отображаем границы
            mdTable.Rows.Alignment = RowAlignment.Center; // центрируем таблицу
            int colCount = table.Rows[0].GetTableCells().Count;
            for (int c = 0; c < colCount; c++)
                mdTable.AddColumn(Unit.FromCentimeter(4)); // фиксированная ширина, можно доработать
            bool isFirstRow = true;
            foreach (var row in table.Rows)
            {
                var mdRow = mdTable.AddRow();
                if (isFirstRow)
                {
                    mdRow.HeadingFormat = true; // первая строка — заголовок
                    isFirstRow = false;
                }
                mdRow.VerticalAlignment = VerticalAlignment.Center; // вертикальное выравнивание
                // mdRow.Height = Unit.FromCentimeter(1); // если нужно фиксировать высоту
                for (int c = 0; c < colCount; c++)
                {
                    var cell = row.GetCell(c);
                    var mdCell = mdRow.Cells[c];
                    if (cell != null)
                    {
                        // Заливка ячейки, если есть цвет
                        var shd = cell.GetCTTc().tcPr?.shd;
                        if (shd != null && !string.IsNullOrEmpty(shd.fill) && shd.fill != "auto" && shd.fill.Length == 6)
                        {
                            try
                            {
                                int rCol = System.Convert.ToInt32(shd.fill.Substring(0, 2), 16);
                                int gCol = System.Convert.ToInt32(shd.fill.Substring(2, 2), 16);
                                int bCol = System.Convert.ToInt32(shd.fill.Substring(4, 2), 16);
                                mdCell.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb((byte)rCol, (byte)gCol, (byte)bCol);
                            }
                            catch { }
                        }
                        foreach (var elem in cell.BodyElements)
                        {
                            if (elem is XWPFParagraph para)
                                ProcessParagraph(para, mdCell);
                            else if (elem is XWPFTable nestedTable)
                                ProcessTable(nestedTable, mdCell);
                        }
                    }
                }
            }
        }

        // Перегрузки для вложенных таблиц/параграфов в ячейках
        private void ProcessParagraph(XWPFParagraph para, Cell cell)
        {
            var mdPara = cell.AddParagraph();
            // (копируем логику ProcessParagraph для Section, но для Cell)
            switch (para.Alignment)
            {
                case ParagraphAlignment.CENTER:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center;
                    break;
                case ParagraphAlignment.RIGHT:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Right;
                    break;
                case ParagraphAlignment.BOTH:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify;
                    break;
                default:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                    break;
            }
            int ToIntSafe(object value)
            {
                if (value == null) return 0;
                if (value is int i) return i;
                if (value is long l) return (int)l;
                if (value is ulong ul) return (int)ul;
                if (value is string s && int.TryParse(s, out int result)) return result;
                return 0;
            }
            int GetSpacingAfterSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.spacing?.after);
            int GetSpacingBeforeSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.spacing?.before);
            int GetIndentationLeftSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.ind?.left);
            int GetIndentationRightSafe(XWPFParagraph p) => ToIntSafe(p.GetCTP()?.pPr?.ind?.right);
            int indentationLeft = GetIndentationLeftSafe(para);
            int indentationRight = GetIndentationRightSafe(para);
            int spacingBefore = GetSpacingBeforeSafe(para);
            int spacingAfter = GetSpacingAfterSafe(para);
            if (indentationLeft > 0)
                mdPara.Format.LeftIndent = indentationLeft / 20.0;
            if (indentationRight > 0)
                mdPara.Format.RightIndent = indentationRight / 20.0;
            if (spacingBefore > 0)
                mdPara.Format.SpaceBefore = spacingBefore / 20.0;
            if (spacingAfter > 0)
                mdPara.Format.SpaceAfter = spacingAfter / 20.0;
            foreach (var run in para.Runs)
            {
                var mdText = mdPara.AddFormattedText(run.Text);
                if (run.IsBold) mdText.Bold = true;
                if (run.IsItalic) mdText.Italic = true;
                if (run.Underline != UnderlinePatterns.None) mdText.Underline = Underline.Single;
                if (run.FontSize > 0) mdText.Size = run.FontSize;
                if (!string.IsNullOrEmpty(run.FontFamily)) mdText.Font.Name = run.FontFamily;
                var clr = run.GetCTR().rPr?.color?.val;
                if (!string.IsNullOrEmpty(clr) && clr.Length == 6)
                {
                    try
                    {
                        int rr = System.Convert.ToInt32(clr.Substring(0, 2), 16);
                        int gg = System.Convert.ToInt32(clr.Substring(2, 2), 16);
                        int bb = System.Convert.ToInt32(clr.Substring(4, 2), 16);
                        mdText.Color = MigraDoc.DocumentObjectModel.Color.FromRgb((byte)rr, (byte)gg, (byte)bb);
                    }
                    catch { }
                }
            }
        }
        private void ProcessTable(XWPFTable table, Cell cell)
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
                                ProcessParagraph(para, mdCell);
                            else if (elem is XWPFTable nestedTable)
                                ProcessTable(nestedTable, mdCell);
                        }
                    }
                }
            }
        }
    }

    public class MigraDocFontResolver : IFontResolver
    {
        private static readonly string FontsFolder = "/Users/ruslanzagidullin/source/docstopdf/XlsxToPdfConverter.Diy/Fonts";

        public byte[] GetFont(string faceName)
        {
            string fontFile = null;
            string lowerFace = faceName.Trim().ToLowerInvariant();
            if (lowerFace.StartsWith("courier new"))
            {
                if (lowerFace.Contains("bi")) fontFile = "courbi.ttf";
                else if (lowerFace.Contains("b")) fontFile = "courbd.ttf";
                else if (lowerFace.Contains("i")) fontFile = "couri.ttf";
                else fontFile = "cour.ttf";
            }
            else
            {
                fontFile = faceName switch
                {
                    "Arial#" => "arial.ttf",
                    "Arial#b" => "arialbd.ttf",
                    "Arial#i" => "ariali.ttf",
                    "Arial#bi" => "arialbi.ttf",
                    "ArialN#" => "ARIALN.TTF",
                    "ArialN#b" => "ARIALNB.TTF",
                    "ArialN#i" => "ARIALNI.TTF",
                    "ArialN#bi" => "ARIALNBI.TTF",
                    "Times#" => "times.ttf",
                    "Times#b" => "timesbd.ttf",
                    "Times#i" => "timesi.ttf",
                    "Times#bi" => "timesbi.ttf",
                    "Calibri#" => "calibri.ttf",
                    "Calibri#b" => "calibrib.ttf",
                    "Calibri#i" => "calibrii.ttf",
                    "Calibri#bi" => "calibriz.ttf",
                    "CambriaMath#" => "CambriaMath.ttf",
                    "ArialBlk#" => "ariblk.ttf",
                    "Courier#" => "cour.ttf",
                    "Courier#b" => "courbd.ttf",
                    "Courier#i" => "couri.ttf",
                    "Courier#bi" => "courbi.ttf",
                    _ => "arial.ttf"
                };
            }
            string path = Path.Combine(FontsFolder, fontFile);
            return File.ReadAllBytes(path);
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            string key = familyName.ToLowerInvariant();
            if (key.Contains("arial"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Arial#bi");
                if (isBold) return new FontResolverInfo("Arial#b");
                if (isItalic) return new FontResolverInfo("Arial#i");
                return new FontResolverInfo("Arial#");
            }
            if (key.Contains("arialn"))
            {
                if (isBold && isItalic) return new FontResolverInfo("ArialN#bi");
                if (isBold) return new FontResolverInfo("ArialN#b");
                if (isItalic) return new FontResolverInfo("ArialN#i");
                return new FontResolverInfo("ArialN#");
            }
            if (key.Contains("times"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Times#bi");
                if (isBold) return new FontResolverInfo("Times#b");
                if (isItalic) return new FontResolverInfo("Times#i");
                return new FontResolverInfo("Times#");
            }
            if (key.Contains("calibri"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Calibri#bi");
                if (isBold) return new FontResolverInfo("Calibri#b");
                if (isItalic) return new FontResolverInfo("Calibri#i");
                return new FontResolverInfo("Calibri#");
            }
            if (key.Contains("cambriamath"))
            {
                return new FontResolverInfo("CambriaMath#");
            }
            if (key.Contains("ariblk"))
            {
                return new FontResolverInfo("ArialBlk#");
            }
            if (key.Contains("courier new") || key.Contains("courier"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Courier New#bi");
                if (isBold) return new FontResolverInfo("Courier New#b");
                if (isItalic) return new FontResolverInfo("Courier New#i");
                return new FontResolverInfo("Courier New#");
            }
            // fallback
            return new FontResolverInfo("Arial#");
        }
    }
} 