using System.Collections.Generic;

namespace XlsxToPdfConverter.Diy
{
    public class FontFamilyModel
    {
        public string Name { get; set; }
        public Dictionary<PdfSharp.Drawing.XFontStyleEx, string> FontFiles { get; set; } = new();
    }
} 