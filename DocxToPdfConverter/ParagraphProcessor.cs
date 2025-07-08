using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.XWPF.UserModel;

namespace DocxToPdfConverter
{
    public static class ParagraphProcessor
    {
        public static void ProcessParagraph(XWPFParagraph para, Section section)
        {
            var mdPara = section.AddParagraph();
            ApplyParagraphFormatting(para, mdPara);
            AddRuns(para, mdPara);
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
                // Добавляем текст, как раньше
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
                        mdText.Color = Color.FromRgb((byte)rr, (byte)gg, (byte)bb);
                    }
                    catch { }
                }

                // --- Добавляем обработку картинок ---
                var pictures = run.GetEmbeddedPictures();
                if (pictures != null && pictures.Count > 0)
                {
                    foreach (var pic in pictures)
                    {
                        var picData = pic.GetPictureData();
                        if (picData != null)
                        {
                            // Определяем расширение файла
                            string ext = picData.SuggestFileExtension();
                            if (string.IsNullOrEmpty(ext)) ext = "png";
                            // Создаем временный файл
                            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "." + ext);
                            File.WriteAllBytes(tempFile, picData.Data);
                            // Вставляем картинку в параграф
                            mdPara.AddImage(tempFile);
                            // Добавляем путь к временному файлу в список для последующего удаления
                            TempImageFiles.Add(tempFile);
                        }
                    }
                }
            }
        }

        // --- Для хранения временных файлов ---
        public static List<string> TempImageFiles = new List<string>();
    }
} 