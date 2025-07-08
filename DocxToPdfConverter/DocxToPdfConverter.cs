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
                foreach (var element in doc.BodyElements)
                {
                    if (element is XWPFParagraph para)
                    {
                        ParagraphProcessor.ProcessParagraph(para, section);
                    }
                    else if (element is XWPFTable table)
                    {
                        TableProcessor.ProcessTable(table, section);
                    }
                }
            }
            var pdfRenderer = new PdfDocumentRenderer(true);
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
    }
} 