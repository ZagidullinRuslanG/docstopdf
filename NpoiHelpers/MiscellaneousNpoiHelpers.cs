using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using SkiaSharp;

namespace NpoiHelpers
{
    public class MiscellaneousNpoiHelpers
    {
        private const int MaxExcelHeightPxPerRow = 546; //409 points. or 546 pixel

        private const int PointsPerInch = 8;

        private const float PxToPtRatio = 1.3333f; //96/72

        /// <summary>
        /// Максимальное значение длины текста в ячейке Excel.
        /// </summary>
        public const short MaxCellTextLenght = short.MaxValue;
        
        /// <summary>
        /// Расчёт минимальной высоты для всей строки.
        /// </summary>
        /// <param name="text">текст для измерения.</param>
        /// <param name="widthPt">ширина ячейки в точках (поинтах). </param>
        /// <param name="currentHeightPx">текущая высота строки, будет использована как минимальное значение.</param>
        /// <param name="fontHeightPt">высота шрифта в точках. Кегль во всем документе = 10, поэтому ставим по-умолчанию.</param>
        /// <param name="fontName">Шрифт Arial по умолчанию.</param>
        /// <param name="maximalMultiply">Сколько строк Excel занимает ячейка (превышения лимита в 409 точек).</param>
        /// <param name="customLineMultiply">Уточнение для высоты строк (для масштабируемых листов, Excel не совсем корректно масштабирует).</param>
        /// <returns>предлагаемая высота в пикселях (можно сразу выставлять в Excel).</returns>
        public static float GetOptimalRowHeightPx(
            string text,
            int widthPt,
            float currentHeightPx = 12,
            int fontHeightPt = 10,
            string fontName = "Arial",
            int maximalMultiply = 1,
            double customLineMultiply = 1)
        {
            if ((int)currentHeightPx >= MaxExcelHeightPxPerRow * maximalMultiply)
            {
                return MaxExcelHeightPxPerRow * maximalMultiply;
            }

            var height = CalcOptimalCellHeightPx(text, fontName, fontHeightPt, widthPt, maximalMultiply, customLineMultiply);

            return (float)( height > currentHeightPx ? height : currentHeightPx);
        }

        /// <summary>
        /// Расчёт минимальной высоты для всей строки.
        /// </summary>
        /// <param name="cellText">Набор тексты для измерения + ширина ячейки в точках (поинтах). </param>
        /// <param name="currentHeightPx">текущая высота строки, будет использована как минимальное значение.</param>
        /// <param name="fontHeightPt">высота шрифта в точках. Кегль во всем документе = 10, поэтому ставим по-умолчанию.</param>
        /// <param name="fontName">Шрифт Arial по умолчанию.</param>
        /// <param name="maximalMultiply">Сколько строк Excel занимает ячейка (превышения лимита в 409 точек).</param>
        /// <param name="customLineMultiply">Уточнение для высоты строк (для масштабируемых листов, Excel не совсем корректно масштабирует).</param>
        /// <returns>предлагаемая высота в пикселях (можно сразу выставлять в Excel).</returns>
        public static float GetOptimalRowHeightPx(
            (string text, int widthPt)[] cellText,
            float currentHeightPx = 12,
            int fontHeightPt = 10,
            string fontName = "Arial",
            int maximalMultiply = 1,
            double customLineMultiply = 1)
        {
            float height = currentHeightPx;
            foreach (var info in cellText)
            {
                height = GetOptimalRowHeightPx(info.text, info.widthPt, height, fontHeightPt,
                    fontName, maximalMultiply, customLineMultiply);
            }

            return height;
        }

        /// <summary>
        /// Расчёт минимальной высоты для ячейки.
        /// </summary>
        /// <param name="text">текст для измерения.</param>
        /// <param name="fontName">Шрифт Arial по умолчанию.</param>
        /// <param name="fontHeightPt">высота шрифта в точках. Кегль во всем документе = 10, поэтому ставим по-умолчанию.</param>
        /// <param name="widthPt">ширина ячейки в точках (поинтах). </param>
        /// <param name="maximalMultiply">Сколько строк Excel занимает ячейка (превышения лимита в 409 точек).</param>
        /// <param name="customLineMultiply">Уточнение для высоты строк (для масштабируемых листов, Excel не совсем корректно масштабирует).</param>
        /// <returns>предлагаемая высота в пикселях (можно сразу выставлять в Excel).</returns>
        public static double CalcOptimalCellHeightPx(
            string text,
            string fontName,
            float fontHeightPt,
            int widthPt,
            int maximalMultiply = 1,
            double customLineMultiply = 1)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0.0;
            }

            using var font = GetSKFont(fontName, fontHeightPt);
            var linesCount = GetLinesCount(widthPt, text, font);

            var metrics = font.Metrics;

            double lineHeight = Math.Ceiling(Math.Abs(metrics.Ascent)) + Math.Ceiling(Math.Abs(metrics.Descent)) + Math.Ceiling(metrics.Leading);

            //компенсируем погрешности изменения масштаба.. выявлено эмпирически..
            lineHeight *= 1.03;

            double result = Math.Min(linesCount * customLineMultiply * lineHeight, MaxExcelHeightPxPerRow * maximalMultiply);
            return result;
        }

        public static SKFont GetSKFont(string fontName, float fontHeightPt)
        {
            var typeface = SKTypeface.FromFamilyName(fontName, SKFontStyle.Bold);
            var font = new SKFont(typeface, fontHeightPt);
            return font;
        }

        public static int GetLinesCount(int widthPt, string text, SKFont font)
        {
            using StringReader textStream = new StringReader(text);
            string line = textStream.ReadLine();
            var cellWidth = widthPt * PointsPerInch;
            double whiteSpaceWidth = font.MeasureText(" ");

            int linesCount = 0;
            while (line != null)
            {
                var regex = new Regex(@"[^\-\s]+|(\-)");
                var matches = regex.Matches(line);
                var wordsInLine = matches.Select(x => x.Value).ToList();

                if (wordsInLine.Count == 0)
                {
                    linesCount++;
                }

                var wordsWidthInLine = new List<double>();

                foreach (var word in wordsInLine)
                {
                    double wordWidth = font.MeasureText(word);
                    wordsWidthInLine.Add(wordWidth * PxToPtRatio);
                }

                double lineWidthSummary = 0;

                for (int i = 0; i < wordsWidthInLine.Count; i++)
                {
                    var wordWithWhiteSpaceWidth = wordsWidthInLine[i] + whiteSpaceWidth;
                    var lineWidth = lineWidthSummary + wordWithWhiteSpaceWidth;

                    if (lineWidth < cellWidth)
                    {
                        lineWidthSummary = lineWidth;
                        if (i == wordsWidthInLine.Count - 1)
                        {
                            linesCount++;
                        }
                    }
                    else
                    {
                        if (wordWithWhiteSpaceWidth >= cellWidth)
                        {
                            linesCount += (int)Math.Ceiling(lineWidth / cellWidth);
                            if (i == wordsWidthInLine.Count - 1)
                            {
                                linesCount++;
                            }
                        }
                        else
                        {
                            linesCount++;
                            i--;
                        }

                        lineWidthSummary = 0;
                    }
                }

                line = textStream.ReadLine();
            }

            var whiteSpacesRegex = new Regex(@"[ ]{2,}");
            var whiteSpacesMatches = whiteSpacesRegex.Matches(text);
            var whiteSpacesInText = whiteSpacesMatches.Select(x => x.Value.Replace("  ", " ")).ToList();
            var whiteSpacesString = string.Join("", whiteSpacesInText);
            double allWhiteSpacesWidth = font.MeasureText(whiteSpacesString) * PxToPtRatio;

            linesCount += (int)( allWhiteSpacesWidth / ( widthPt * PointsPerInch));
            return linesCount;
        }

        public static void SetPageSetupA3(IWorkbook workbook, ISheet sheet, int sheetNumber, int rowNumber, int colNumber)
        {
            workbook.SetPrintArea(sheetNumber, 0, colNumber, 0, rowNumber);
            sheet.SetMargin(MarginType.RightMargin, 0.1d);
            sheet.SetMargin(MarginType.TopMargin, 0.1d);
            sheet.SetMargin(MarginType.LeftMargin, 0.1d);
            sheet.SetMargin(MarginType.BottomMargin, 0.1d);
            sheet.PrintSetup.PaperSize = (short)PaperSize.A3 + 1;
        }

        public static void SetPageSetupA4(IWorkbook workbook, ISheet sheet, int sheetNumber, int rowNumber, int colNumber)
        {
            workbook.SetPrintArea(sheetNumber, 0, colNumber, 0, rowNumber);
            sheet.SetMargin(MarginType.RightMargin, 0.1d);
            sheet.SetMargin(MarginType.TopMargin, 0.1d);
            sheet.SetMargin(MarginType.LeftMargin, 0.1d);
            sheet.SetMargin(MarginType.BottomMargin, 0.1d);
            sheet.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;
        }
        
        /// <summary>
        /// Установка границ печати листа Excel (с настройкой количества страниц на листе).
        /// </summary>
        /// <param name="workbook">Книга Excel.</param>
        /// <param name="sheet">Лист.</param>
        /// <param name="sheetNumber">Номер листа.</param>
        /// <param name="rowNumber">Номер граничной строки.</param>
        /// <param name="colNumber">Номер граничного столбца.</param>
        /// <param name="fitWidth">Число страниц в ширину.</param>
        /// <param name="fitHeight">Число страниц в высоту. Если 0 - то значение будет "Авто"</param>
        public static void SetPageSetupA4FitToPage(IWorkbook workbook, ISheet sheet, int sheetNumber, int rowNumber, int colNumber, short fitWidth, short fitHeight)
        {
            SetPageSetupA4(workbook, sheet, sheetNumber, rowNumber, colNumber);
            sheet.FitToPage = true; // заставляет вписывать документ в необходимые размеры при печати.
            sheet.PrintSetup.FitWidth = fitWidth; // задает размер по ширине в страницах.
            sheet.PrintSetup.FitHeight = fitHeight; // задает размер по высоте в страницах. Если 0 - то значение будет "Авто".
        }

        public static IRichTextString CreateAndFormatText(string text, IFont baseFont, IFont emFont, List<int> emInds)
        {
            IRichTextString format = new XSSFRichTextString(text.ToString());

            format.ApplyFont(baseFont);

            for (int i = 0; i < emInds.Count - 2; i += 2)
            {
                format.ApplyFont(emInds[i], emInds[i + 1] + 1, emFont);
            }

            format.ApplyFont(emInds[emInds.Count - 2], emInds[emInds.Count - 1] + 1, emFont);

            return format;
        }

        /// <summary>
        /// Преобразование количества пикселей в ширину (столбца).
        /// </summary>
        /// <param name="pixel">Количество пикселей.</param>
        public static int Px2width(int pixel) => ( pixel * 256) / 7;
    }
}
