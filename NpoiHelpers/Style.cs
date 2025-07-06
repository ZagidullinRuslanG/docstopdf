using NPOI.SS.UserModel;

namespace NpoiHelpers
{
    /// <summary>
    /// Определяет стиль оформления фрагмента текста.
    /// </summary>
    public class Style
    {
        public int StartIndex { get; }

        public int Length { get; }

        public IFont FontStyle { get; }

        public Style(int startIndex, int length, IFont fontStyle)
        {
            StartIndex = startIndex;
            Length = length;
            FontStyle = fontStyle;
        }
    }
}