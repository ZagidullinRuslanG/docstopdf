using System.Text;
using System.Text.RegularExpressions;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using DocxModel = NPOI.XWPF.UserModel;
using MdModel = MigraDoc.DocumentObjectModel;

namespace DocxToPdfConverter
{
    public class Bookmark
    {
        public string Name { get; init; }
        public string Title { get; init; }
        public string Number { get; init; }
        public bool? IsLandscape { get; set; }
        public int Order { get; init; }
        public bool UsedForPageSettings { get; set; }
        public bool UsedForParagraph { get; set; }
        public PageFields PageFields { get; set; }
    }

    public record PageFields(Unit Top, Unit Bottom, Unit Left, Unit Right);

    public static class ParagraphProcessor
    {
        public static void ProcessParagraph(XWPFParagraph para, HeaderFooter footer)
        {
            var mdPara = footer.AddParagraph();
            ApplyParagraphFormatting(para, mdPara);
            AddRuns(para, mdPara, string.Empty);
        }

        public static void ProcessParagraph(
            XWPFParagraph para,
            Section section,
            Dictionary<string, List<Bookmark>> bookmarks,
            bool isNewSection)
        {
            if ((para.IsPageBreak || HasVisualBreak(para)) && !isNewSection && !string.IsNullOrWhiteSpace(para.Text) )
            {
                section.AddPageBreak();
            }

            var mdPara = section.AddParagraph();

            var prefix = string.Empty;
            if (!string.IsNullOrWhiteSpace(para.Style))
            {
                var styleName = GetStyle(para.Document, para.StyleID).Name.ToLower();
                if (styleName.StartsWith("heading") && !string.IsNullOrWhiteSpace(para.Text))
                {
                    var title = para.Text.Trim();

                    var bookmarkName = title;
                    var bookmarkDuplicates = bookmarks.GetValueOrDefault(bookmarkName);
                    var bookmark = bookmarkDuplicates?.Where(b => !b.UsedForParagraph).MinBy(b => b.Order);

                    if (bookmark is not null)
                    {
                        bookmarkName += bookmark.Order;
                        mdPara.AddBookmark(bookmarkName);
                        prefix = bookmark.Number;
                        bookmark.UsedForParagraph = true;
                    }
                }
            }
            ApplyParagraphFormatting(para, mdPara);

            AddRuns(para, mdPara, prefix);
        }

        public static void ProcessSdt(XWPFSDT sdt, Section section, List<Bookmark> bookmarks)
        {
            if (sdt.ElementType != BodyElementType.CONTENTCONTROL)
            {
                return;
            }

            var usableWidth = Unit.FromMillimeter(170);
            var doc = section.Document;
            if (doc is not null)
            {
                var pageWidth = section.PageSetup.PageWidth.Point == 0 ?  doc.DefaultPageSetup.PageWidth : section.PageSetup.PageWidth;
                var leftMargin = section.PageSetup.LeftMargin.Point == 0 ? doc.DefaultPageSetup.LeftMargin : section.PageSetup.LeftMargin;
                var rightMargin = section.PageSetup.RightMargin.Point == 0 ?  doc.DefaultPageSetup.RightMargin : section.PageSetup.RightMargin;
                usableWidth = pageWidth - leftMargin - rightMargin;
            }

            foreach (var bookmark in bookmarks.OrderBy(b => b.Order))
            {
                var paragraph = section.AddParagraph();
                paragraph.Format.SpaceAfter = Unit.Zero;

                const int maxPageNumberDigits = 3;
                const double charWidthInPoints = 7.5;
                const double pageNumberWidth = maxPageNumberDigits * charWidthInPoints;

                var pageUsableWidth = usableWidth.Point;
                var rightTabPos = pageUsableWidth - 2;
                var fillerPos = rightTabPos - pageNumberWidth;

                // Настраиваем табуляции
                paragraph.Format.TabStops.Clear();
                paragraph.Format.TabStops.AddTabStop(fillerPos, TabAlignment.Left, TabLeader.Dots);

                var bookmarkName = bookmark.Name + bookmark.Order;

                var hyperlink = paragraph.AddHyperlink(bookmarkName);
                hyperlink.AddText($"{bookmark.Title}");
                paragraph.AddTab();
                paragraph.AddPageRefField(bookmarkName);
            }
        }

        //todo: поворот текста в ячейке

        private static void ApplyParagraphFormatting(XWPFParagraph para, Paragraph mdPara)
        {
            //сначала применить форматирование из стиля
            ApplyStyleParagraphFormatting(para.Document, para.StyleID, mdPara);

            //применить явное форматирование параграфа
            ApplyNativeParagraphFormattingFromStyle(para, mdPara);
        }

        private static void ApplyNativeParagraphFormattingFromStyle(XWPFParagraph para, Paragraph mdPara)
        {
            mdPara.Format.Alignment = para.Alignment switch
            {
                DocxModel.ParagraphAlignment.CENTER => MdModel.ParagraphAlignment.Center,
                DocxModel.ParagraphAlignment.RIGHT => MdModel.ParagraphAlignment.Right,
                DocxModel.ParagraphAlignment.BOTH => MdModel.ParagraphAlignment.Justify,
                _ => mdPara.Format.Alignment
            };

            var pPr = para.GetCTP()?.pPr;
            SetIndents(pPr, mdPara.Format);

            //Шрифт
            var rPr = pPr?.rPr;
            var rFonts = rPr?.rFonts; //шрифт
            if (rFonts is not null)
            {
                mdPara.Format.Font.Name = rFonts.ascii;
                mdPara.Format.Font.Bold = rPr.b?.val ?? false;
                mdPara.Format.Font.Italic = rPr.i?.val ?? false;

                mdPara.Format.Font.Underline =
                    (rPr.u?.val ?? ST_Underline.none) != ST_Underline.none ? Underline.Single : Underline.None;
            }

            var size = rPr?.sz; //размер
            if (size is not null)
            {
                mdPara.Format.Font.Size = Unit.FromPoint(size.val / 2);
            }
        }

        private static void SetIndents(CT_PPr pPr, ParagraphFormat format)
        {
            if (pPr?.ind?.left is not null && int.TryParse(pPr?.ind?.left, out var left))
            {
                format.LeftIndent = Unit.FromPoint(left / 20.0);
            }
            if (pPr?.ind?.right is not null && int.TryParse(pPr?.ind?.right, out var right))
            {
                format.RightIndent = Unit.FromPoint(right / 20.0);
            }
            if (pPr?.ind?.hanging is not null)
            {
                format.FirstLineIndent = Unit.FromPoint( -(long)pPr.ind.hanging / 20.0);
            }

            var before = pPr?.spacing?.before;
            if (before is not null)
            {
                format.SpaceBefore = Unit.FromPoint(before.Value / 20.0);
            }

            var after = pPr?.spacing?.after;
            if (after is not null)
            {
                format.SpaceAfter = Unit.FromPoint(after.Value / 20.0);
            }
        }

        private static void ApplyStyleParagraphFormatting(XWPFDocument doc, string styleId, Paragraph mdPara)
        {
            var style = GetStyle(doc, styleId);

            var ctStyle = style?.GetCTStyle();
            if (ctStyle is null)
            {
                return;
            }

            var parentStyle = ctStyle.basedOn?.val; //Для рекурсивного обхода базовых стилей
            if (parentStyle is not null)
            {
                ApplyStyleParagraphFormatting(doc, parentStyle, mdPara);
            }

            //Выравнивание
            var ctStylePPr = ctStyle.pPr;   //настройки параграфа в стиле
            var ctJc = ctStylePPr?.jc?.val; //выравнивание
            mdPara.Format.Alignment = ctJc switch
            {
                ST_Jc.center => MdModel.ParagraphAlignment.Center,
                ST_Jc.right => MdModel.ParagraphAlignment.Right,
                ST_Jc.both => MdModel.ParagraphAlignment.Justify,
                _ => mdPara.Format.Alignment
            };

            //Отступы
            SetIndents(ctStylePPr, mdPara.Format);

            var ctStyleRPr = ctStyle.rPr; //настройки run в стиле

            //Шрифт
            var rFonts = ctStyleRPr?.rFonts; //шрифт
            if (rFonts is not null)
            {
                mdPara.Format.Font.Name = rFonts.ascii;
                mdPara.Format.Font.Bold = ctStyleRPr.b?.val ?? false;
                mdPara.Format.Font.Italic = ctStyleRPr.i?.val ?? false;

                mdPara.Format.Font.Underline =
                    (ctStyleRPr.u?.val ?? ST_Underline.none) != ST_Underline.none ? Underline.Single : Underline.None;
            }

            var size = ctStyleRPr?.sz; //размер
            if (size is not null)
            {
                mdPara.Format.Font.Size = Unit.FromPoint(size.val / 2);
            }
        }

        private static bool HasVisualBreak(XWPFParagraph paragraph)
        {
            var rList = paragraph.GetCTP().GetRList();

            return rList
                .Any(
                    run => run.ItemsElementName
                        .Any(n => n == RunItemsChoiceType.lastRenderedPageBreak));
        }

        public static void ProcessParagraph(XWPFParagraph para, Cell cell, int maxWordLen)
        {
            var mdPara = cell.AddParagraph();
            ApplyParagraphFormatting(para, mdPara);
            mdPara.Format.SpaceAfter = Unit.Zero;
            AddRuns(para, mdPara, string.Empty, maxWordLen);
        }

        private static string InsertSoftSpaces(string text, int interval)
        {
            var result = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                result.Append(text[i]);
                if (i % interval == 0 && i > 0)
                    result.Append('\u00AD'); //мягкий перенос
            }
            return result.ToString();
        }

        internal static XWPFStyle GetStyle(XWPFDocument doc, string styleId)
        {
            return string.IsNullOrWhiteSpace(styleId) ? null : doc.GetStyles().GetStyle(styleId);
        }

        private static void AddRuns(XWPFParagraph para, Paragraph mdPara, string prefix, int maxWordLength = 0)
        {
            foreach (var run in para.Runs)
            {
                // 1. Добавляем текст, если есть
                if (!string.IsNullOrEmpty(run.Text))
                {
                    var text = run.Text;

                    text = ApplyMaxLength(text, maxWordLength);

                    if (!string.IsNullOrWhiteSpace(prefix))
                    {
                        text = $"{prefix} {run.Text}";
                    }
                    var mdText = mdPara.AddFormattedText(text);
                    prefix = string.Empty;

                    var ctR = run.GetCTR();
                    var rPr = ctR.rPr;

                    if (rPr is not null)
                    {
                        if (rPr.b?.val ?? false)
                            mdText.Bold = true;
                        if (rPr.i?.val ?? false)
                            mdText.Italic = true;
                        if (rPr.u is not null && rPr.u?.val != ST_Underline.none)
                            mdText.Underline = Underline.Single;
                        if (rPr.sz?.val > 0)
                            mdText.Size = rPr.sz.val / 2;
                        if (!string.IsNullOrEmpty(rPr.rFonts?.ascii))
                            mdText.Font.Name = rPr.rFonts.ascii;

                        var clr = rPr.color?.val;
                        if (!string.IsNullOrEmpty(clr) && clr.Length == 6)
                            mdText.Color = Color.FromRgb(
                                (byte)int.Parse(clr[..2], System.Globalization.NumberStyles.HexNumber),
                                (byte)int.Parse(clr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                                (byte)int.Parse(clr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber));
                    }

                }

                // 2. Добавляем картинки inline
                var pictures = run.GetEmbeddedPictures();
                if (pictures is { Count: > 0 })
                {
                    foreach (var pic in pictures)
                    {
                        var picData = pic.GetPictureData();
                        if (picData is { Data.Length: > 0 })
                        {
                            string base64 = Convert.ToBase64String(picData.Data);
                            string imgStr = "base64:" + base64;
                            var image = mdPara.AddImage(imgStr);
                            image.LockAspectRatio = true;
                            image.Width = Unit.FromPoint(pic.Width / 12700);
                            image.Height = Unit.FromPoint(pic.Height / 12700);
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

        private static string ApplyMaxLength(string text, int maxWordLength)
        {
            if (maxWordLength == 0)
            {
                return text;
            }

            var words = text.Split(" ");

            var resWords = words
                .Select(word => word.Length < maxWordLength ? word : InsertSoftSpaces(word, maxWordLength));

            return string.Join(" ", resWords);
        }

        public static Bookmark FindBookmark(
            XWPFParagraph para,
            List<int> numeratorState,
            Dictionary<string, List<Bookmark>> bookmarks,
            ref int bookmarkOrder)
        {
            if (!string.IsNullOrWhiteSpace(para.Text) && !string.IsNullOrWhiteSpace(para.Style))
            {
                //вытащить нумерацию заголовков
                var style = GetStyle(para.Document, para.StyleID);
                var styleName = style.Name.ToLower();
                if (styleName.StartsWith("heading"))
                {
                    var number = string.Empty;

                    var regex = new Regex(@"\d+");
                    var match = regex.Match(styleName);

                    if (match.Success && int.TryParse(match.Value, out int level)
                        && HasNumber(para, style))
                    {
                        while (numeratorState.Count < level)
                        {
                            numeratorState.Add(0);
                        }

                        for (int i = level + 1; i < numeratorState.Count - 1; i++)
                        {
                            numeratorState[i] = 0;
                        }

                        numeratorState[level - 1]++;

                        number = string.Join('.', numeratorState.Take(level));
                    }

                    var name = para.Text.Trim();

                    if (!bookmarks.TryGetValue(name, out var duplicates))
                    {
                        duplicates = new List<Bookmark>();
                        bookmarks.Add(name, duplicates);
                    }

                    var bookmark = new Bookmark
                    {
                        Title = $"{number} {name}".Trim(),
                        Number = number,
                        Name = name,
                        Order = bookmarkOrder++,
                    };

                    duplicates.Add(bookmark);
                    return bookmark;
                }
            }

            return null;
        }

        private static bool HasNumber(XWPFParagraph para, XWPFStyle style)
        {
            var paragraphNumId = para.GetNumID();

            if (paragraphNumId is null)
            {
                var pPrNumPr = style.GetCTStyle().pPr.numPr;

                if (pPrNumPr?.numId != null)
                {
                    return pPrNumPr.numId.val != "0";
                }

                return false;
            }

            return paragraphNumId != "0";
        }

        // --- Для хранения временных файлов ---
        public static readonly List<string> TempImageFiles = new();
    }
}