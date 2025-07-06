using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiHelpers
{
    public class StyleList
    {
        private List<Style> styles;
        private IFont defaultFont;

        /// <summary>
        /// ВАЖНО: баг NPOI, касающиеся стили (например "te[/style1][style2]xt") обрабатываются некорректно, style1 не будет применен.
        /// Должен быть хотя бы один символ между двумя стилями.
        /// </summary>
        public void AddStyle(int startIndex, int length, IFont fontStyle)
        {
            styles.Add(new Style(startIndex, length, fontStyle));
        }

        public StyleList()
        {
            styles = new List<Style>();
        }

        public StyleList(IFont defaultFont)
        {
            this.defaultFont = defaultFont;
            styles = new List<Style>();
        }

        public void SetDefaultFont(IFont defaultFont)
        {
            this.defaultFont = defaultFont;
        }

        public IFont GetDefaultFont()
        {
            return defaultFont;
        }

        public IRichTextString ApplyStyles(string cellLine)
        {
            IRichTextString cellRichText = new XSSFRichTextString(cellLine);
            return ApplyStyles(cellRichText);
        }

        public IRichTextString ApplyStyles(StringBuilder cellLine)
        {
            IRichTextString cellRichText = new XSSFRichTextString(cellLine.ToString());
            return ApplyStyles(cellRichText);
        }

        public IRichTextString ApplyStyles(IRichTextString cellRichText)
        {
            if (defaultFont != null)
            {
                cellRichText.ApplyFont(defaultFont);
            }

            foreach (var style in styles)
            {
                cellRichText.ApplyFont(style.StartIndex, style.StartIndex + style.Length, style.FontStyle);
            }

            return cellRichText;
        }

        private IFont FindStyle(int position)
        {
            return styles.Find(item => item.StartIndex <= position && position < item.StartIndex + item.Length)?.FontStyle
                   ?? defaultFont;
        }

        /// <summary>
        /// "Применяет" список стилей на строку текста: разбивает на части и группирует с соответствующим стилем текста.
        /// </summary>
        /// <param name="source">Входная строка.</param>
        public List<StyleTextPair> SplitAsApplied(string source)
        {
            var separators = new List<int> { 0 };
            var output = new List<StyleTextPair>();

            styles.Sort((a, b) => a.StartIndex - b.StartIndex);
            int last = 0;
            foreach (var style in styles)
            {
                int current = style.StartIndex;
                if (current != last)
                {
                    separators.Add(current);
                }

                last = current + style.Length;
                separators.Add(last);
            }

            if (last != source.Length)
            {
                separators.Add(source.Length);
            }

            for (int i = 1; i < separators.Count; i++)
            {
                output.Add(new StyleTextPair(
                    source.Substring(separators[i - 1], separators[i] - separators[i - 1]),
                    FindStyle(separators[i - 1])));
            }

            return output;
        }
    }
}