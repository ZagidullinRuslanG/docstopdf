using PdfSharp.Drawing;
using NPOI.OpenXmlFormats.Wordprocessing;
using PdfSharp.Fonts;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using NPOI.XWPF.UserModel;
using Document = MigraDoc.DocumentObjectModel.Document;
using ParagraphAlignment = NPOI.XWPF.UserModel.ParagraphAlignment;
using System.Reflection;

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
                foreach (var para in doc.Paragraphs)
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
                    int GetSpacingAfterSafe(XWPFParagraph para)
                    {
                        var spacing = para.GetCTP()?.pPr?.spacing;
                        return ToIntSafe(spacing?.after);
                    }
                    int GetSpacingBeforeSafe(XWPFParagraph para)
                    {
                        var spacing = para.GetCTP()?.pPr?.spacing;
                        return ToIntSafe(spacing?.before);
                    }
                    int GetIndentationLeftSafe(XWPFParagraph para)
                    {
                        var ind = para.GetCTP()?.pPr?.ind;
                        return ToIntSafe(ind?.left);
                    }
                    int GetIndentationRightSafe(XWPFParagraph para)
                    {
                        var ind = para.GetCTP()?.pPr?.ind;
                        return ToIntSafe(ind?.right);
                    }
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
            }
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.Save(pdfPath);
        }

        // Заготовка для отрисовки таблицы через NPOI
        private void DrawTableNpoi(XWPFTable table, XGraphics gfx, double left, ref double y, double tableWidth)
        {
            int rowCount = table.NumberOfRows;
            if (rowCount == 0) return;
            int colCount = table.Rows[0].GetTableCells().Count;
            if (colCount == 0) return;
            double[] colWidths = new double[colCount];
            double totalWidth = 0;
            for (int c = 0; c < colCount; c++)
            {
                var tcW = table.Rows[0].GetCell(c)?.GetCTTc().tcPr?.tcW;
                if (tcW != null && int.TryParse(tcW.w, out int w) && w > 0)
                {
                    colWidths[c] = w / 20.0; // docx: twentieths of a point
                }
                else
                {
                    colWidths[c] = tableWidth / colCount;
                }
                totalWidth += colWidths[c];
            }
            // Нормализация ширины
            for (int c = 0; c < colCount; c++)
                colWidths[c] = colWidths[c] / totalWidth * tableWidth;
            XFont defaultFont = new XFont("Arial", 12);
            bool[,] occupied = new bool[rowCount, colCount];
            // Считаем высоту каждой строки по максимальному количеству строк в ячейках
            double[] rowHeights = new double[rowCount];
            for (int r = 0; r < rowCount; r++)
            {
                double maxLines = 1;
                var row = table.Rows[r];
                for (int c = 0; c < colCount; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null) continue;
                    int lines = 0;
                    foreach (var para in cell.Paragraphs)
                    {
                        var paraLines = (para.Text ?? "").Split('\n');
                        lines += Math.Max(paraLines.Length, 1);
                    }
                    if (lines > maxLines) maxLines = lines;
                }
                rowHeights[r] = maxLines * defaultFont.Height + 4;
            }
            double yPos = y;
            for (int r = 0; r < rowCount; r++)
            {
                var row = table.Rows[r];
                double x = left;
                for (int c = 0; c < colCount; c++)
                {
                    if (occupied[r, c]) { x += colWidths[c]; continue; }
                    var cell = row.GetCell(c);
                    if (cell == null) { x += colWidths[c]; continue; }
                    int colspan = 1;
                    int rowspan = 1;
                    var gridSpan = cell.GetCTTc().tcPr?.gridSpan;
                    if (gridSpan != null && int.TryParse(gridSpan.val, out int gs) && gs > 1)
                        colspan = gs;
                    var vMerge = cell.GetCTTc().tcPr?.vMerge;
                    if (vMerge != null && vMerge.val == ST_Merge.@continue)
                    {
                        x += colWidths[c];
                        continue;
                    }
                    double mergedWidth = colWidths[c] * colspan;
                    double mergedHeight = 0;
                    for (int dr = 0; dr < rowspan; dr++)
                        if (r + dr < rowCount)
                            mergedHeight += rowHeights[r + dr];
                    for (int dr = 0; dr < rowspan; dr++)
                        for (int dc = 0; dc < colspan; dc++)
                            if (r + dr < rowCount && c + dc < colCount)
                                occupied[r + dr, c + dc] = true;
                    var shd = cell.GetCTTc().tcPr?.shd;
                    if (shd != null && !string.IsNullOrEmpty(shd.fill) && shd.fill != "auto")
                    {
                        try
                        {
                            int rCol = System.Convert.ToInt32(shd.fill.Substring(0, 2), 16);
                            int gCol = System.Convert.ToInt32(shd.fill.Substring(2, 2), 16);
                            int bCol = System.Convert.ToInt32(shd.fill.Substring(4, 2), 16);
                            gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(rCol, gCol, bCol)), x, yPos, mergedWidth, mergedHeight);
                        }
                        catch { }
                    }
                    gfx.DrawRectangle(XPens.Black, x, yPos, mergedWidth, mergedHeight);
                    double textY = yPos + 2;
                    foreach (var para in cell.Paragraphs)
                    {
                        foreach (var run in para.Runs)
                        {
                            string[] lines = (run.Text ?? "").Split('\n');
                            // Определяем стиль для XFont
                            XFont font;
                            if (run.IsBold)
                                font = new XFont("Arial", 12, XFontStyleEx.Bold);
                            else if (run.IsItalic)
                                font = new XFont("Arial", 12, XFontStyleEx.Italic);
                            else
                                font = new XFont("Arial", 12, XFontStyleEx.Regular);
                            XBrush brush = XBrushes.Black;
                            if (run.GetCTR().rPr?.color != null && !string.IsNullOrEmpty(run.GetCTR().rPr.color.val))
                            {
                                var clr = run.GetCTR().rPr.color.val;
                                if (clr.Length == 6)
                                {
                                    try
                                    {
                                        int rr = System.Convert.ToInt32(clr.Substring(0, 2), 16);
                                        int gg = System.Convert.ToInt32(clr.Substring(2, 2), 16);
                                        int bb = System.Convert.ToInt32(clr.Substring(4, 2), 16);
                                        brush = new XSolidBrush(XColor.FromArgb(rr, gg, bb));
                                    }
                                    catch { }
                                }
                            }
                            foreach (var line in lines)
                            {
                                // Автоматический перенос текста по ширине ячейки
                                var rect = new XRect(x + 2, textY, mergedWidth - 4, font.Height * 2);
                                var format = new XStringFormat { LineAlignment = XLineAlignment.Near };
                                var tf = new PdfSharp.Drawing.Layout.XTextFormatter(gfx);
                                // Определяем выравнивание
                                var align = para.Alignment;
                                switch (align)
                                {
                                    case ParagraphAlignment.CENTER:
                                        tf.Alignment = PdfSharp.Drawing.Layout.XParagraphAlignment.Center;
                                        break;
                                    case ParagraphAlignment.RIGHT:
                                        tf.Alignment = PdfSharp.Drawing.Layout.XParagraphAlignment.Right;
                                        break;
                                    default:
                                        tf.Alignment = PdfSharp.Drawing.Layout.XParagraphAlignment.Left;
                                        break;
                                }
                                tf.DrawString(line, font, brush, rect, format);
                                textY += font.Height;
                            }
                            // Вставка картинок из run
                            foreach (var pic in run.GetEmbeddedPictures())
                            {
                                var imgData = pic.GetPictureData();
                                if (imgData != null)
                                {
                                    using (var ms = new System.IO.MemoryStream(imgData.Data))
                                    {
                                        try
                                        {
                                            var ximg = XImage.FromStream(ms);
                                            double imgW = Math.Min(ximg.PixelWidth * 0.75, mergedWidth - 4);
                                            double imgH = ximg.PixelHeight * 0.75;
                                            gfx.DrawImage(ximg, x + 2, textY, imgW, imgH);
                                            textY += imgH + 2;
                                        }
                                        catch { }
                                    }
                                }
                            }
                            // Вложенные таблицы
                            foreach (var bodyElem in cell.BodyElements)
                            {
                                if (bodyElem is XWPFTable nestedTable)
                                {
                                    double nestedY = textY;
                                    DrawTableNpoi(nestedTable, gfx, x + 2, ref nestedY, mergedWidth - 4);
                                    textY = nestedY;
                                }
                            }
                        }
                    }
                    x += mergedWidth;
                }
                yPos += rowHeights[r];
            }
            y = yPos;
        }
        // Заготовка для отрисовки картинки через NPOI
        private void DrawPictureNpoi(XWPFPictureData pic, XGraphics gfx, double left, ref double y)
        {
            // TODO: реализовать отрисовку картинки через NPOI
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