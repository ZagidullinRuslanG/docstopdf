using NPOI.SS.UserModel;

namespace NpoiHelpers
{
    public class StyleTextPair
    {
        public string Text;
        public IFont Font;

        public StyleTextPair(string text, IFont font)
        {
            this.Text = text;
            this.Font = font;
        }
    }
}