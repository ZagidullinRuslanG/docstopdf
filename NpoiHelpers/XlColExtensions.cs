using System;

namespace NpoiHelpers
{
    public static class XlColExtensions
    {
        // Преобразует букву столбца Excel (например, "A", "AB") в номер (0-индексация)
        public static int XlCol(this string col)
        {
            if (string.IsNullOrEmpty(col)) throw new ArgumentException("col");
            int result = 0;
            foreach (char c in col.ToUpper())
            {
                if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column letter");
                result = result * 26 + (c - 'A' + 1);
            }
            return result - 1;
        }

        // Преобразует номер столбца (0-индексация) в букву Excel (например, 0 -> "A", 27 -> "AB")
        public static string XlCol(this int col)
        {
            if (col < 0) throw new ArgumentException("col");
            string result = string.Empty;
            col++;
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }
    }
} 