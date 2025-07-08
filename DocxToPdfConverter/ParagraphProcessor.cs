using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;

namespace DocxToPdfConverter
{
    public static class ParagraphProcessor
    {
        public static void ProcessParagraph(XWPFParagraph para, DocumentObject docObj)
        {
            Paragraph mdPara;

            switch (docObj)
            {
                case Section section:
                {
                    if (para.IsPageBreak || HasVisualBreak(para))
                    {
                        section.AddPageBreak();
                    }

                    mdPara = section.AddParagraph();
                    break;
                }
                case HeaderFooter footer:
                    mdPara = footer.AddParagraph();
                    break;
                default:
                    return;
            }

            ApplyParagraphFormatting(para, mdPara);
            AddRuns(para, mdPara);
        }

        private static bool HasVisualBreak(XWPFParagraph paragraph)
        {
            var rList = paragraph.GetCTP().GetRList();

            return rList
                .Any(run => run.ItemsElementName
                    .Any(n => n == RunItemsChoiceType.lastRenderedPageBreak));
        }

        public static void ProcessParagraph(XWPFParagraph para, Cell cell)
        {
            var mdPara = cell.AddParagraph();
            ApplyParagraphFormatting(para, mdPara);
            AddRuns(para, mdPara);
        }

        private static void ApplyParagraphFormatting(XWPFParagraph para, Paragraph mdPara)
        {
            switch (para.Alignment)
            {
                case NPOI.XWPF.UserModel.ParagraphAlignment.CENTER:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center;
                    break;
                case NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Right;
                    break;
                case NPOI.XWPF.UserModel.ParagraphAlignment.BOTH:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify;
                    break;
                default:
                    mdPara.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                    break;
            }
            int indentationLeft = DocxValueUtils.GetIndentationLeftSafe(para);
            int indentationRight = DocxValueUtils.GetIndentationRightSafe(para);
            int spacingBefore = DocxValueUtils.GetSpacingBeforeSafe(para);
            int spacingAfter = DocxValueUtils.GetSpacingAfterSafe(para);
            if (indentationLeft > 0)
                mdPara.Format.LeftIndent = indentationLeft / 20.0;
            if (indentationRight > 0)
                mdPara.Format.RightIndent = indentationRight / 20.0;
            if (spacingBefore > 0)
                mdPara.Format.SpaceBefore = spacingBefore / 20.0;
            if (spacingAfter > 0)
                mdPara.Format.SpaceAfter = spacingAfter / 20.0;
        }

        private static void AddRuns(XWPFParagraph para, Paragraph mdPara)
        {
            foreach (var run in para.Runs)
            {
                // 1. Добавляем текст, если есть
                if (!string.IsNullOrEmpty(run.Text))
                {
                    var mdText = mdPara.AddFormattedText(run.Text);
                    if (run.IsBold) mdText.Bold = true;
                    if (run.IsItalic) mdText.Italic = true;
                    if (run.Underline != UnderlinePatterns.None) mdText.Underline = Underline.Single;
                    if (run.FontSize > 0) mdText.Size = run.FontSize;
                    if (!string.IsNullOrEmpty(run.FontFamily)) mdText.Font.Name = run.FontFamily;
                    var clr = run.GetCTR().rPr?.color?.val;
                    if (!string.IsNullOrEmpty(clr) && clr.Length == 6)
                        mdText.Color = Color.FromRgb(
                            (byte)int.Parse(clr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber),
                            (byte)int.Parse(clr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                            (byte)int.Parse(clr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber));
                }
                // 2. Добавляем картинки inline
                var pictures = run.GetEmbeddedPictures();
                if (pictures != null && pictures.Count > 0)
                {
                    foreach (var pic in pictures)
                    {
                        var picData = pic.GetPictureData();
                        if (picData != null && picData.Data != null && picData.Data.Length > 0)
                        {
                            string base64 = Convert.ToBase64String(picData.Data);
                            string imgStr = "base64:" + base64;
                            // Логирование информации о картинке
                            var image = mdPara.AddImage(imgStr);
                            image.LockAspectRatio = true;
                            image.Width = "5cm";
                        }
                        else
                        {
                            Console.WriteLine("[Image] Warning: Empty or null image data");
                        }
                    }
                }
            }
            // Если нет контента, не добавлять пустой параграф (для таблиц)
        }

        // --- Для хранения временных файлов ---
        public static List<string> TempImageFiles = new List<string>();
    }
} 