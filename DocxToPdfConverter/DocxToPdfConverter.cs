using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Linq;
using PdfSharp.Fonts;
using XlsxToPdfConverter.Diy;
using DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;

namespace DocxToPdfConverter
{
    public class DocxToPdfConverter
    {
        static DocxToPdfConverter()
        {
            GlobalFontSettings.FontResolver = new CustomFontResolver();
        }

        private class CellInfo
        {
            public DocumentFormat.OpenXml.Wordprocessing.TableCell Cell;
            public int Row;
            public int Col;
            public int RowSpan = 1;
            public int ColSpan = 1;
            public bool IsMerged = false;
        }

        public void Convert(string docxPath, string pdfPath)
        {
            using (PdfDocument pdf = new PdfDocument())
            {
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                double y = 40;
                double left = 40;
                double right = page.Width - XUnit.FromPoint(40);
                double lineHeight = 18;
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var numbering = wordDoc.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
                    // Найти первый параграф, если он styleId == "Title" или похожий, считать его заголовком таблицы
                    var firstPara = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                    if (firstPara != null)
                    {
                        string titleText = string.Concat(firstPara.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().Select(r => string.Concat(r.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text))));
                        if (!string.IsNullOrWhiteSpace(titleText))
                        {
                            XFont titleFont = new XFont("Times New Roman", 16, XFontStyleEx.Italic);
                            double maxWidth = right - left;
                            int titleLines = EstimateLineCount(titleText, titleFont, maxWidth, gfx);
                            double titleLineHeight = gfx.MeasureString("A", titleFont).Height;
                            double titleHeight = titleLines * titleLineHeight + 8;
                            var tf = new XTextFormatterEx2(gfx);
                            tf.DrawString(titleText, titleFont, XBrushes.Black, new XRect(XUnit.FromPoint(left), XUnit.FromPoint(y), XUnit.FromPoint(maxWidth), XUnit.FromPoint(titleHeight)), XStringFormats.TopLeft);
                            y += titleHeight + 8;
                            // Удалить этот параграф из body, чтобы не рисовать его второй раз
                            firstPara.Remove();
                        }
                    }
                    foreach (var element in body.Elements())
                    {
                        if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
                        {
                            string styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
                            double curLineHeight = GetLineHeightForStyle(styleId);
                            double x = left;
                            var indent = para.ParagraphProperties?.Indentation;
                            double leftIndent = 0, rightIndent = 0, firstLineIndent = 0;
                            if (indent != null)
                            {
                                if (indent.Left != null && double.TryParse(indent.Left.Value, out double l))
                                    leftIndent = l / 20.0;
                                if (indent.Right != null && double.TryParse(indent.Right.Value, out double r))
                                    rightIndent = r / 20.0;
                                if (indent.FirstLine != null && double.TryParse(indent.FirstLine.Value, out double f))
                                    firstLineIndent = f / 20.0;
                            }
                            int listLevel = 0;
                            string marker = null;
                            var numProp = para.ParagraphProperties?.NumberingProperties;
                            if (numProp != null)
                            {
                                if (numProp.NumberingLevelReference?.Val != null)
                                    listLevel = (int)numProp.NumberingLevelReference.Val.Value;
                                marker = GetListMarker(numProp, numbering);
                            }
                            double listIndent = listLevel * 20;
                            x += leftIndent + firstLineIndent + listIndent;
                            double paragraphWidth = right - left - leftIndent - rightIndent - firstLineIndent - listIndent;
                            var justification = para.ParagraphProperties?.Justification?.Val;
                            XStringFormat stringFormat = XStringFormats.TopLeft;
                            if (justification != null)
                            {
                                switch (justification.Value)
                                {
                                    case JustificationValues.Center:
                                        stringFormat = XStringFormats.TopCenter;
                                        break;
                                    case JustificationValues.Right:
                                        stringFormat = XStringFormats.TopRight;
                                        break;
                                    case JustificationValues.Both:
                                        stringFormat = XStringFormats.TopLeft;
                                        break;
                                    default:
                                        stringFormat = XStringFormats.TopLeft;
                                        break;
                                }
                            }
                            if (!string.IsNullOrEmpty(marker))
                            {
                                var markerFont = new XFont("Arial", 12, XFontStyleEx.Regular);
                                gfx.DrawString(marker, markerFont, XBrushes.Black, new XRect(XUnit.FromPoint(x), XUnit.FromPoint(y), XUnit.FromPoint(30), XUnit.FromPoint(curLineHeight)), stringFormat);
                                x += gfx.MeasureString(marker, markerFont).Width + 5;
                            }
                            // Обработка разрывов страниц
                            bool pageBreak = false;
                            foreach (var br in para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Break>())
                            {
                                if (br.Type != null && br.Type.Value == DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page)
                                {
                                    pageBreak = true;
                                    break;
                                }
                            }
                            if (pageBreak)
                            {
                                page = pdf.AddPage();
                                gfx.Dispose();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 40;
                                continue;
                            }
                            foreach (var run in para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                            {
                                string text = string.Concat(run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                                if (string.IsNullOrEmpty(text)) continue;
                                var runProps = run.RunProperties;
                                XFontStyleEx style = XFontStyleEx.Regular;
                                if (runProps != null)
                                {
                                    if (runProps.Bold != null) style |= XFontStyleEx.Bold;
                                    if (runProps.Italic != null) style |= XFontStyleEx.Italic;
                                    if (runProps.Underline != null) style |= XFontStyleEx.Underline;
                                }
                                double fontSize = 12;
                                if (runProps?.FontSize?.Val != null)
                                {
                                    if (double.TryParse(runProps.FontSize.Val.Value, out double sz))
                                        fontSize = sz / 2.0;
                                }
                                string fontName = runProps?.RunFonts?.Ascii?.Value ?? "Arial";
                                var runFont = new XFont(fontName, fontSize, style);
                                XBrush brush = XBrushes.Black;
                                if (runProps?.Color?.Val != null)
                                {
                                    var color = runProps.Color.Val.Value;
                                    if (color.Length == 6)
                                    {
                                        int r = System.Convert.ToInt32(color.Substring(0, 2), 16);
                                        int g = System.Convert.ToInt32(color.Substring(2, 2), 16);
                                        int b = System.Convert.ToInt32(color.Substring(4, 2), 16);
                                        brush = new XSolidBrush(XColor.FromArgb(r, g, b));
                                    }
                                }
                                gfx.DrawString(text, runFont, brush, new XRect(XUnit.FromPoint(x), XUnit.FromPoint(y), XUnit.FromPoint(paragraphWidth), XUnit.FromPoint(curLineHeight)), stringFormat);
                                x += gfx.MeasureString(text, runFont).Width;
                            }
                            // Вставка изображений (Drawing)
                            foreach (var drawing in para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
                            {
                                var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                                if (blip != null)
                                {
                                    var relId = blip.Embed.Value;
                                    var imgPart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(relId);
                                    using (var imgStream = imgPart.GetStream())
                                    using (var ms = new MemoryStream())
                                    {
                                        imgStream.CopyTo(ms);
                                        ms.Position = 0;
                                        XImage ximg = XImage.FromStream(new MemoryStream(ms.ToArray()));
                                        // Определяем размеры (по умолчанию 100x100pt)
                                        double imgWidth = 100, imgHeight = 100;
                                        var ext = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
                                        if (ext != null)
                                        {
                                            imgWidth = ext.Cx / 12700.0; // EMU to pt
                                            imgHeight = ext.Cy / 12700.0;
                                        }
                                        gfx.DrawImage(ximg, XUnit.FromPoint(x), XUnit.FromPoint(y), XUnit.FromPoint(imgWidth), XUnit.FromPoint(imgHeight));
                                        y += imgHeight + 5; // перенос строки после картинки
                                    }
                                }
                            }
                            y += curLineHeight;
                            if (y > page.Height - XUnit.FromPoint(40))
                            {
                                page = pdf.AddPage();
                                gfx.Dispose();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 40;
                            }
                        }
                        else if (element is DocumentFormat.OpenXml.Wordprocessing.Table table)
                        {
                            y = DrawTable(table, gfx, left, y, right - left, wordDoc);
                            if (y > page.Height - XUnit.FromPoint(40))
                            {
                                page = pdf.AddPage();
                                gfx.Dispose();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 40;
                            }
                        }
                    }
                }
                gfx.Dispose();
                pdf.Save(pdfPath);
            }
        }

        private string GetListMarker(NumberingProperties numProp, Numbering numbering)
        {
            if (numProp.NumberingId == null) return null;
            string marker = "• ";
            if (numbering != null && numProp.NumberingId != null)
            {
                var numId = numProp.NumberingId.Val.Value;
                var abstractNumId = numbering.Elements<NumberingInstance>()
                    .FirstOrDefault(n => n.NumberID.Value == numId)?
                    .AbstractNumId?.Val?.Value;
                if (abstractNumId != null)
                {
                    var abstractNum = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.AbstractNum>()
                        .FirstOrDefault(n => n.AbstractNumberId.Value == abstractNumId);
                    if (abstractNum != null)
                    {
                        var lvl = abstractNum.Elements<Level>().FirstOrDefault();
                        if (lvl != null && lvl.NumberingFormat != null)
                        {
                            if (lvl.NumberingFormat.Val == NumberFormatValues.Bullet)
                                marker = "• ";
                            else if (lvl.NumberingFormat.Val == NumberFormatValues.Decimal)
                                marker = "1. ";
                        }
                    }
                }
            }
            return marker;
        }

        private double DrawTable(DocumentFormat.OpenXml.Wordprocessing.Table table, XGraphics gfx, double left, double y, double tableWidth, WordprocessingDocument wordDoc)
        {
            double cellPadding = 4;
            // Строим виртуальную матрицу ячеек
            var rows = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList();
            int rowCount = rows.Count;
            int colCount = 0;
            foreach (var row in rows)
            {
                int count = 0;
                foreach (var cell in row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                {
                    int gridSpan = cell.TableCellProperties?.GridSpan?.Val != null ? (int)cell.TableCellProperties.GridSpan.Val : 1;
                    count += gridSpan;
                }
                if (count > colCount) colCount = count;
            }
            CellInfo[,] cellMatrix = new CellInfo[rowCount, colCount];
            for (int row = 0; row < rowCount; row++)
            {
                int col = 0;
                foreach (var cell in rows[row].Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                {
                    // Пропустить уже занятые позиции (поглощённые)
                    while (col < colCount && cellMatrix[row, col] != null) col++;
                    int colspan = cell.TableCellProperties?.GridSpan?.Val != null ? (int)cell.TableCellProperties.GridSpan.Val : 1;
                    // Определяем rowspan (VerticalMerge)
                    int rowspan = 1;
                    var vMerge = cell.TableCellProperties?.VerticalMerge;
                    if (vMerge != null && vMerge.Val != null && vMerge.Val.Value == DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Restart)
                    {
                        for (int r = row + 1; r < rowCount; r++)
                        {
                            var nextRowCells = rows[r].Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();
                            int nextColIdx = 0;
                            foreach (var nextCell in nextRowCells)
                            {
                                int nextGridSpan = nextCell.TableCellProperties?.GridSpan?.Val != null ? (int)nextCell.TableCellProperties.GridSpan.Val : 1;
                                if (nextColIdx == col)
                                {
                                    var nextVMerge = nextCell.TableCellProperties?.VerticalMerge;
                                    if (nextVMerge != null && nextVMerge.Val == null)
                                    {
                                        rowspan++;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                    break;
                                }
                                nextColIdx += nextGridSpan;
                            }
                        }
                    }
                    for (int dr = 0; dr < rowspan; dr++)
                    {
                        for (int dc = 0; dc < colspan; dc++)
                        {
                            if ((row + dr) >= rowCount || (col + dc) >= colCount)
                                continue; // Не выходить за пределы массива
                            cellMatrix[row + dr, col + dc] = new CellInfo
                            {
                                Cell = cell,
                                Row = row,
                                Col = col,
                                RowSpan = rowspan,
                                ColSpan = colspan,
                                IsMerged = !(dr == 0 && dc == 0)
                            };
                        }
                    }
                    col += colspan;
                }
            }
            double cellWidth = tableWidth / colCount;
            double curY = y;
            XFont defaultFont = new XFont("Arial", 12, XFontStyleEx.Regular);
            double[] rowHeights = new double[rowCount];
            for (int row = 0; row < rowCount; row++)
            {
                double maxCellHeight = 0;
                for (int col = 0; col < colCount; col++)
                {
                    if (cellMatrix[row, col] == null || cellMatrix[row, col].IsMerged) continue;
                    var cell = cellMatrix[row, col].Cell;
                    int colspan = cellMatrix[row, col].ColSpan;
                    double availableWidth = cellWidth * colspan - 2 * cellPadding;
                    double cellTextHeight = 0;
                    foreach (var para in cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        foreach (var run in para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                        {
                            string text = string.Concat(run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                            if (string.IsNullOrEmpty(text)) continue;
                            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProps = run.RunProperties;
                            XFontStyleEx style = XFontStyleEx.Regular;
                            if (runProps != null)
                            {
                                if (runProps.Bold != null) style |= XFontStyleEx.Bold;
                                if (runProps.Italic != null) style |= XFontStyleEx.Italic;
                                if (runProps.Underline != null) style |= XFontStyleEx.Underline;
                            }
                            double fontSize = 12;
                            if (runProps?.FontSize?.Val != null)
                            {
                                if (double.TryParse(runProps.FontSize.Val.Value, out double sz))
                                    fontSize = sz / 2.0;
                            }
                            string fontName = runProps?.RunFonts?.Ascii?.Value ?? "Arial";
                            var runFont = new XFont(fontName, fontSize, style);
                            var words = text.Split(' ');
                            string currentLine = "";
                            foreach (var word in words)
                            {
                                string testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                                var size = gfx.MeasureString(testLine, runFont);
                                if (size.Width > availableWidth)
                                {
                                    if (!string.IsNullOrEmpty(currentLine))
                                        cellTextHeight += gfx.MeasureString(currentLine, runFont).Height;
                                    currentLine = word;
                                }
                                else
                                {
                                    currentLine = testLine;
                                }
                            }
                            if (!string.IsNullOrEmpty(currentLine))
                                cellTextHeight += gfx.MeasureString(currentLine, runFont).Height;
                        }
                        cellTextHeight += gfx.MeasureString("A", new XFont("Arial", 12)).Height * 0.5;
                    }
                    cellTextHeight += 2 * cellPadding;
                    if (cellTextHeight > maxCellHeight) maxCellHeight = cellTextHeight;
                }
                rowHeights[row] = maxCellHeight > 0 ? maxCellHeight : 18;
            }
            for (int row = 0; row < rowCount; row++)
            {
                double x = left;
                for (int col = 0; col < colCount; col++)
                {
                    var info = cellMatrix[row, col];
                    if (info == null || info.IsMerged) continue;
                    var cell = info.Cell;
                    int colspan = info.ColSpan;
                    int rowspan = info.RowSpan;
                    double availableWidth = cellWidth * colspan - 2 * cellPadding;
                    // Высота объединённой ячейки — сумма rowHeights по covered строкам
                    double cellHeight = 0;
                    for (int r = row; r < row + rowspan; r++)
                        cellHeight += rowHeights[r];
                    // Получаем параметры границ
                    var borders = cell.TableCellProperties?.TableCellBorders;
                    XPen penTop = GetCellBorderPen(borders?.TopBorder);
                    XPen penBottom = GetCellBorderPen(borders?.BottomBorder);
                    XPen penLeft = GetCellBorderPen(borders?.LeftBorder);
                    XPen penRight = GetCellBorderPen(borders?.RightBorder);
                    // Рисуем каждую сторону отдельно
                    gfx.DrawLine(penTop, XUnit.FromPoint(x), XUnit.FromPoint(curY), XUnit.FromPoint(x + cellWidth * colspan), XUnit.FromPoint(curY));
                    gfx.DrawLine(penBottom, XUnit.FromPoint(x), XUnit.FromPoint(curY + cellHeight), XUnit.FromPoint(x + cellWidth * colspan), XUnit.FromPoint(curY + cellHeight));
                    gfx.DrawLine(penLeft, XUnit.FromPoint(x), XUnit.FromPoint(curY), XUnit.FromPoint(x), XUnit.FromPoint(curY + cellHeight));
                    gfx.DrawLine(penRight, XUnit.FromPoint(x + cellWidth * colspan), XUnit.FromPoint(curY), XUnit.FromPoint(x + cellWidth * colspan), XUnit.FromPoint(curY + cellHeight));
                    // Определяем выравнивание
                    var vAlign = cell.TableCellProperties?.TableCellVerticalAlignment?.Val;
                    XStringFormat stringFormat = XStringFormats.TopLeft;
                    if (vAlign != null)
                    {
                        switch (vAlign.Value)
                        {
                            case DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues.Center:
                                stringFormat = XStringFormats.CenterLeft;
                                break;
                            case DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues.Bottom:
                                stringFormat = XStringFormats.BottomLeft;
                                break;
                            default:
                                stringFormat = XStringFormats.TopLeft;
                                break;
                        }
                    }
                    // Горизонтальное выравнивание по первому параграфу
                    var para = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                    if (para != null)
                    {
                        var justification = para.ParagraphProperties?.Justification?.Val;
                        if (justification != null)
                        {
                            switch (justification.Value)
                            {
                                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center:
                                    stringFormat = XStringFormats.TopCenter;
                                    break;
                                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right:
                                    stringFormat = XStringFormats.TopRight;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    // Рисуем содержимое ячейки (текст, вложенные таблицы, картинки)
                    double textX = x + cellPadding;
                    double textY = curY + cellPadding;
                    foreach (var element in cell.Elements())
                    {
                        if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph p)
                        {
                            foreach (var run in p.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                            {
                                string text = string.Concat(run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                                if (string.IsNullOrEmpty(text)) continue;
                                DocumentFormat.OpenXml.Wordprocessing.RunProperties runProps = run.RunProperties;
                                XFontStyleEx style = XFontStyleEx.Regular;
                                if (runProps != null)
                                {
                                    if (runProps.Bold != null) style |= XFontStyleEx.Bold;
                                    if (runProps.Italic != null) style |= XFontStyleEx.Italic;
                                    if (runProps.Underline != null) style |= XFontStyleEx.Underline;
                                }
                                double fontSize = 12;
                                if (runProps?.FontSize?.Val != null)
                                {
                                    if (double.TryParse(runProps.FontSize.Val.Value, out double sz))
                                        fontSize = sz / 2.0;
                                }
                                string fontName = runProps?.RunFonts?.Ascii?.Value ?? "Arial";
                                var runFont = new XFont(fontName, fontSize, style);
                                XBrush brush = XBrushes.Black;
                                if (runProps?.Color?.Val != null)
                                {
                                    var color = runProps.Color.Val.Value;
                                    if (color.Length == 6)
                                    {
                                        int r = System.Convert.ToInt32(color.Substring(0, 2), 16);
                                        int g = System.Convert.ToInt32(color.Substring(2, 2), 16);
                                        int b = System.Convert.ToInt32(color.Substring(4, 2), 16);
                                        brush = new XSolidBrush(XColor.FromArgb(r, g, b));
                                    }
                                }
                                gfx.DrawString(text, runFont, brush, new XRect(XUnit.FromPoint(textX), XUnit.FromPoint(textY), XUnit.FromPoint(availableWidth), XUnit.FromPoint(cellHeight - 2 * cellPadding)), stringFormat);
                                textY += gfx.MeasureString(text, runFont).Height;
                            }
                            // Вставка изображений (Drawing) внутри параграфа
                            foreach (var drawing in p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
                            {
                                var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                                if (blip != null)
                                {
                                    var relId = blip.Embed.Value;
                                    var imgPart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(relId);
                                    using (var imgStream = imgPart.GetStream())
                                    using (var ms = new MemoryStream())
                                    {
                                        imgStream.CopyTo(ms);
                                        ms.Position = 0;
                                        XImage ximg = XImage.FromStream(new MemoryStream(ms.ToArray()));
                                        double imgWidth = 100, imgHeight = 100;
                                        var ext = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
                                        if (ext != null)
                                        {
                                            imgWidth = ext.Cx / 12700.0;
                                            imgHeight = ext.Cy / 12700.0;
                                        }
                                        gfx.DrawImage(ximg, XUnit.FromPoint(textX), XUnit.FromPoint(textY), XUnit.FromPoint(imgWidth), XUnit.FromPoint(imgHeight));
                                        textY += imgHeight + 5;
                                    }
                                }
                            }
                        }
                        else if (element is DocumentFormat.OpenXml.Wordprocessing.Table nestedTable)
                        {
                            // Вложенная таблица
                            textY = DrawTable(nestedTable, gfx, textX, textY, availableWidth, wordDoc);
                        }
                    }
                    x += cellWidth * colspan;
                }
                curY += rowHeights[row];
            }
            return curY + 8;
        }

        private XFont GetFontForStyle(string styleId)
        {
            if (styleId == "Heading1") return new XFont("Arial", 20, XFontStyleEx.Bold);
            if (styleId == "Heading2") return new XFont("Arial", 16, XFontStyleEx.Bold);
            if (styleId == "Heading3") return new XFont("Arial", 14, XFontStyleEx.Bold);
            return new XFont("Arial", 12, XFontStyleEx.Regular);
        }
        private double GetLineHeightForStyle(string styleId)
        {
            if (styleId == "Heading1") return 28;
            if (styleId == "Heading2") return 22;
            if (styleId == "Heading3") return 18;
            return 16;
        }
        private XFont GetFontForRun(XFont baseFont, DocumentFormat.OpenXml.Wordprocessing.RunProperties runProps)
        {
            XFontStyleEx style = XFontStyleEx.Regular;
            if (runProps != null)
            {
                if (runProps.Bold != null) style |= XFontStyleEx.Bold;
                if (runProps.Italic != null) style |= XFontStyleEx.Italic;
                if (runProps.Underline != null) style |= XFontStyleEx.Underline;
            }
            return new XFont(baseFont.FontFamily.Name, baseFont.Size, style);
        }

        // Функция для word wrap и подсчёта строк:
        int EstimateLineCount(string text, XFont font, double maxWidth, XGraphics gfx)
        {
            if (string.IsNullOrEmpty(text)) return 1;
            var words = text.Split(' ');
            int lines = 1;
            string currentLine = "";
            foreach (var word in words)
            {
                string testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                var size = gfx.MeasureString(testLine, font);
                if (size.Width > maxWidth)
                {
                    lines++;
                    currentLine = word;
                }
                else
                {
                    currentLine = testLine;
                }
            }
            return lines;
        }

        // Вспомогательная функция для получения XPen по границе ячейки
        private XPen GetCellBorderPen(DocumentFormat.OpenXml.Wordprocessing.BorderType border)
        {
            if (border == null || border.Val == null || border.Val.Value == DocumentFormat.OpenXml.Wordprocessing.BorderValues.Nil)
                return XPens.Transparent;
            double width = 1.0;
            if (border.Size != null)
                width = (double)border.Size.Value / 8.0; // OpenXML size в 1/8 pt
            XColor color = XColors.Black;
            if (border.Color != null && border.Color.Value != "auto")
            {
                var c = border.Color.Value;
                if (c.Length == 6)
                {
                    int r = System.Convert.ToInt32(c.Substring(0, 2), 16);
                    int g = System.Convert.ToInt32(c.Substring(2, 2), 16);
                    int b = System.Convert.ToInt32(c.Substring(4, 2), 16);
                    color = XColor.FromArgb(r, g, b);
                }
            }
            XPen pen = new XPen(color, width);
            if (border.Val != null)
            {
                switch (border.Val.Value)
                {
                    case DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted:
                        pen.DashStyle = XDashStyle.Dot;
                        break;
                    case DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed:
                        pen.DashStyle = XDashStyle.Dash;
                        break;
                    default:
                        pen.DashStyle = XDashStyle.Solid;
                        break;
                }
            }
            return pen;
        }
    }
} 