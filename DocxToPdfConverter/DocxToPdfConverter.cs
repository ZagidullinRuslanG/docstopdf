using MigraDoc.DocumentObjectModel;
using PdfSharp.Fonts;
using MigraDoc.Rendering;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using Document = MigraDoc.DocumentObjectModel.Document;
using MdModel = MigraDoc.DocumentObjectModel;

namespace DocxToPdfConverter
{
    public class DocxToPdfConverter
    {
        static DocxToPdfConverter()
        {
            GlobalFontSettings.FontResolver = new MigraDocFontResolver();
        }

        private static void ApplyDefaultSettings(Document migraDoc)
        {
            var normalStyle = migraDoc.Styles[StyleNames.Normal]
                ?? migraDoc.Styles.AddStyle(StyleNames.Normal, string.Empty);

            normalStyle.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);
            normalStyle.Font.Size = Unit.FromPoint(12);
        }

        private static void ApplyMainStyleFormatting(XWPFDocument doc, Document migraDoc)
        {
            var npoiNormalStyle = doc.GetStyles().GetStyleWithName("Normal"); // или стиль по умолчанию

            var migraNormalStyle = migraDoc.Styles[StyleNames.Normal]
                ?? migraDoc.Styles.AddStyle(StyleNames.Normal, string.Empty);

            if (npoiNormalStyle?.GetCTStyle()?.pPr?.ind != null)
            {
                var npoiIndent = npoiNormalStyle.GetCTStyle().pPr.ind;

                if (double.TryParse(npoiIndent.left, out var dLeft))
                {
                    migraNormalStyle.ParagraphFormat.LeftIndent = Unit.FromPoint(dLeft / 20.0);
                }

                if (double.TryParse(npoiIndent.right, out var dRight))
                {
                    migraNormalStyle.ParagraphFormat.LeftIndent = Unit.FromPoint(dRight / 20.0);
                }

                migraNormalStyle.ParagraphFormat.FirstLineIndent = Unit.FromPoint(npoiIndent.hanging / 20.0);
                migraNormalStyle.ParagraphFormat.FirstLineIndent = Unit.FromPoint(npoiIndent.firstLine / 20.0);
            }

            // Если есть настройки интервалов
            if (npoiNormalStyle?.GetCTStyle()?.pPr?.spacing != null)
            {
                var npoiSpacing = npoiNormalStyle.GetCTStyle().pPr.spacing;

                if (npoiSpacing.before != null)
                    migraNormalStyle.ParagraphFormat.SpaceBefore = Unit.FromPoint(npoiSpacing.before.Value / 20.0);

                if (npoiSpacing.after != null)
                    migraNormalStyle.ParagraphFormat.SpaceAfter = Unit.FromPoint(npoiSpacing.after.Value / 20.0);

                if (double.TryParse(npoiSpacing.line, out var dLine))
                {
                    migraNormalStyle.ParagraphFormat.LineSpacingRule = MdModel.LineSpacingRule.AtLeast;
                    migraNormalStyle.ParagraphFormat.LineSpacing = Unit.FromPoint(dLine / 20.0);
                }
            }
        }

        public static void Convert(string docxPath, string pdfPath)
        {
            var document = new Document();

            using (var stream = File.OpenRead(docxPath))
            {
                var doc = new XWPFDocument(stream);

                ApplyDefaultSettings(document);
                ApplyMainStyleFormatting(doc, document);

                var bookmarks = CollectBookmarks(doc);
                ProcessBodyElements(doc, document, bookmarks);

                //колонтитул первой страницы
                var firstPageFooter = doc.GetHeaderFooterPolicy().GetFirstPageFooter();
                if (firstPageFooter is not null)
                {
                    var section = document.Sections.OfType<Section>().FirstOrDefault();
                    if (section is not null)
                    {
                        section.PageSetup.DifferentFirstPageHeaderFooter = true;

                        var firstPageSection = section.Footers.FirstPage;
                        ProcessBodyElements(firstPageFooter, firstPageSection!);
                    }
                }

                //колонтитулы остальных страниц
                var defaultFooter = doc.GetHeaderFooterPolicy().GetDefaultFooter();
                if (defaultFooter is not null)
                {
                    foreach (var section in document.Sections.OfType<Section>())
                    {
                        var evenPageSection = section.Footers.Primary;
                        ProcessBodyElements(defaultFooter, evenPageSection!);
                    }
                }
            }

            var pdfRenderer = new PdfDocumentRenderer { Document = document };
            pdfRenderer.RenderDocument();
            pdfRenderer.Save(pdfPath);

            // Удаляем временные файлы картинок
            foreach (var tempFile in ParagraphProcessor.TempImageFiles)
            {
                try
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                }
                catch
                {
                    // ignored
                }
            }

            ParagraphProcessor.TempImageFiles.Clear();
        }

        private static void ProcessBodyElements(IBody body, Document document, Dictionary<string, List<Bookmark>> bookmarks)
        {
            Section section = null;
            var list = bookmarks.SelectMany(b => b.Value).ToList();

            foreach (var element in body.BodyElements)
            {
                switch (element)
                {
                    case XWPFParagraph para:
                        section = BuildSectionIfNecessary(document, para, section, bookmarks, out var isNewSection);
                        ParagraphProcessor.ProcessParagraph(para, section, bookmarks, isNewSection);
                        break;

                    case XWPFSDT sdt:
                        section ??= BuildSection(document);

                        ParagraphProcessor.ProcessSdt(sdt, section, list);
                        break;

                    case XWPFTable table:
                        section ??= BuildSection(document);
                        TableProcessor.ProcessTable(table, section);
                        break;
                }
            }
        }

        private static Section BuildSection(Document document)
        {
            var section = document.AddSection();
            
            section.PageSetup.TopMargin = Unit.FromCentimeter(1.75);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(2);

            section.PageSetup.LeftMargin = Unit.FromCentimeter(2);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1);

            return section;
        }

        private static Section BuildSectionIfNecessary(
            Document document,
            XWPFParagraph para,
            Section section,
            Dictionary<string, List<Bookmark>> bookmarks,
            out bool isNew)
        {
            isNew = section is null;

            if (string.IsNullOrWhiteSpace(para.Style))
            {
                return section ?? BuildSection(document);
            }

            var styleName = ParagraphProcessor.GetStyle(para.Document, para.StyleID).Name.ToLower();
            if (styleName.StartsWith("heading"))
            {
                var title = para.Text.Trim();
                var bookmarkDuplicates = bookmarks.GetValueOrDefault(title);
                var bookmark = bookmarkDuplicates?.Where(b => !b.UsedForPageSettings).MinBy(b => b.Order);

                if (bookmark is not null)
                {
                    bookmark.UsedForPageSettings = true;

                    var setLandscape = bookmark.IsLandscape != null
                        && bookmark.IsLandscape != (section?.PageSetup.Orientation == Orientation.Landscape);

                    var setFields = bookmark.PageFields != null && (
                        section?.PageSetup.BottomMargin != bookmark.PageFields.Bottom ||
                        section?.PageSetup.RightMargin != bookmark.PageFields.Right ||
                        section?.PageSetup.LeftMargin != bookmark.PageFields.Left ||
                        section?.PageSetup.TopMargin != bookmark.PageFields.Top);

                    var buildSection = setLandscape || setFields;

                    if (buildSection)
                    {
                        section = BuildSection(document);

                        if (setLandscape)
                        {
                            section.PageSetup.Orientation =
                                bookmark.IsLandscape.Value ? Orientation.Landscape : Orientation.Portrait;
                        }

                        if (setFields)
                        {
                            section.PageSetup.LeftMargin = bookmark.PageFields.Left;
                            section.PageSetup.RightMargin = bookmark.PageFields.Right;
                            section.PageSetup.TopMargin = bookmark.PageFields.Top;
                            section.PageSetup.BottomMargin = bookmark.PageFields.Bottom;
                        }

                        isNew = true;
                        return section;
                    }
                }
            }

            return section ?? BuildSection(document);
        }

        private static Dictionary<string, List<Bookmark>> CollectBookmarks(IBody doc)
        {
            var numeratorState = new List<int>();

            var bookmarkMap = new Dictionary<string, List<Bookmark>>();
            Bookmark currentBookmark = null;
            var order = 0;

            foreach (var element in doc.BodyElements)
            {
                if (element is XWPFParagraph p)
                {
                    var bookmark = ParagraphProcessor.FindBookmark(p, numeratorState, bookmarkMap, ref order);
                    if (bookmark is not null)
                    {
                        currentBookmark = bookmark;
                    }

                    if (currentBookmark is not null)
                    {
                        var orientationInsideOfBlock = FindOrientation(p);
                        if (orientationInsideOfBlock is not null)
                        {
                            currentBookmark.IsLandscape = orientationInsideOfBlock.Value;
                        }

                        var pageFields = FindPageFields(p);
                        if (pageFields is not null)
                        {
                            currentBookmark.PageFields = new PageFields(
                                Unit.FromPoint(pageFields.top / 20),
                                Unit.FromPoint(pageFields.bottom / 20),
                                Unit.FromPoint(pageFields.left / 20),
                                Unit.FromPoint(pageFields.right / 20));
                        }
                    }
                }
            }

            return bookmarkMap;
        }

        private static bool? FindOrientation(XWPFParagraph para)
        {
            var pPrSectPr = para.GetCTP().pPr.sectPr;

            if (pPrSectPr is null)
            {
                return null;
            }

            var orient = pPrSectPr.pgSz.orient;

            return orient == ST_PageOrientation.landscape;
        }

        private static CT_PageMar FindPageFields(XWPFParagraph para)
        {
            var pPrSectPr = para.GetCTP().pPr.sectPr;

            return pPrSectPr?.pgMar;
        }

        private static void ProcessBodyElements(IBody body, HeaderFooter footer)
        {
            foreach (var element in body.BodyElements)
            {
                switch (element)
                {
                    case XWPFParagraph para:
                        ParagraphProcessor.ProcessParagraph(para, footer);
                        break;

                    case XWPFTable table:
                        TableProcessor.ProcessTable(table, footer);
                        break;
                }
            }
        }
    }
}