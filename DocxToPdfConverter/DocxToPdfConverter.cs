using MigraDoc.DocumentObjectModel;
using PdfSharp.Fonts;
using MigraDoc.Rendering;
using NPOI.XWPF.UserModel;
using Document = MigraDoc.DocumentObjectModel.Document;

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

                ProcessBodyElements(doc, section);

                //колонтитул первой страницы
                var firstPageFooter = doc.GetHeaderFooterPolicy().GetFirstPageFooter();
                if (firstPageFooter is not null)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = true;

                    var firstPageSection = section.Footers.FirstPage;
                    ProcessBodyElements(firstPageFooter, firstPageSection!);
                }

                //колонтитулы остальных страниц
                var evenPageFooter = doc.GetHeaderFooterPolicy().GetDefaultFooter();
                if (evenPageFooter is not null)
                {
                    var evenPageSection = section.Footers.Primary;
                    ProcessBodyElements(evenPageFooter, evenPageSection!);
                }


            }
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.Save(pdfPath);

            // Удаляем временные файлы картинок
            foreach (var tempFile in ParagraphProcessor.TempImageFiles)
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
            ParagraphProcessor.TempImageFiles.Clear();
        }

        private static void ProcessBodyElements(IBody doc, DocumentObject section)
        {
            foreach (var element in doc.BodyElements)
            {
                switch (element)
                {
                    case XWPFParagraph para:
                        ParagraphProcessor.ProcessParagraph(para, section);
                        break;

                    case XWPFTable table:
                        TableProcessor.ProcessTable(table, section);
                        break;
                }
            }
        }
    }
} 