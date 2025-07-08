using PdfSharp.Fonts;

namespace DocxToPdfConverter
{
    public class MigraDocFontResolver : IFontResolver
    {
        private static readonly string FontsFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");

        public byte[] GetFont(string faceName)
        {
            string fontFile = null;
            string lowerFace = faceName.Trim().ToLowerInvariant();
            if (lowerFace.StartsWith("courier new"))
            {
                if (lowerFace.Contains("bi")) fontFile = "courbi.ttf";
                else if (lowerFace.Contains("b")) fontFile = "courbd.ttf";
                else if (lowerFace.Contains("i")) fontFile = "couri.ttf";
                else fontFile = "cour.ttf";
            }
            else
            {
                fontFile = faceName switch
                {
                    "Arial#" => "arial.ttf",
                    "Arial#b" => "arialbd.ttf",
                    "Arial#i" => "ariali.ttf",
                    "Arial#bi" => "arialbi.ttf",
                    "ArialN#" => "ARIALN.TTF",
                    "ArialN#b" => "ARIALNB.TTF",
                    "ArialN#i" => "ARIALNI.TTF",
                    "ArialN#bi" => "ARIALNBI.TTF",
                    "Times#" => "times.ttf",
                    "Times#b" => "timesbd.ttf",
                    "Times#i" => "timesi.ttf",
                    "Times#bi" => "timesbi.ttf",
                    "Calibri#" => "calibri.ttf",
                    "Calibri#b" => "calibrib.ttf",
                    "Calibri#i" => "calibrii.ttf",
                    "Calibri#bi" => "calibriz.ttf",
                    "CambriaMath#" => "CambriaMath.ttf",
                    "ArialBlk#" => "ariblk.ttf",
                    "Courier#" => "cour.ttf",
                    "Courier#b" => "courbd.ttf",
                    "Courier#i" => "couri.ttf",
                    "Courier#bi" => "courbi.ttf",
                    _ => "arial.ttf"
                };
            }
            string path = Path.Combine(FontsFolder, fontFile);
            return File.ReadAllBytes(path);
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            string key = familyName.ToLowerInvariant();
            if (key.Contains("arial"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Arial#bi");
                if (isBold) return new FontResolverInfo("Arial#b");
                if (isItalic) return new FontResolverInfo("Arial#i");
                return new FontResolverInfo("Arial#");
            }
            if (key.Contains("arialn"))
            {
                if (isBold && isItalic) return new FontResolverInfo("ArialN#bi");
                if (isBold) return new FontResolverInfo("ArialN#b");
                if (isItalic) return new FontResolverInfo("ArialN#i");
                return new FontResolverInfo("ArialN#");
            }
            if (key.Contains("times"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Times#bi");
                if (isBold) return new FontResolverInfo("Times#b");
                if (isItalic) return new FontResolverInfo("Times#i");
                return new FontResolverInfo("Times#");
            }
            if (key.Contains("calibri"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Calibri#bi");
                if (isBold) return new FontResolverInfo("Calibri#b");
                if (isItalic) return new FontResolverInfo("Calibri#i");
                return new FontResolverInfo("Calibri#");
            }
            if (key.Contains("cambriamath"))
            {
                return new FontResolverInfo("CambriaMath#");
            }
            if (key.Contains("ariblk"))
            {
                return new FontResolverInfo("ArialBlk#");
            }
            if (key.Contains("courier new") || key.Contains("courier"))
            {
                if (isBold && isItalic) return new FontResolverInfo("Courier New#bi");
                if (isBold) return new FontResolverInfo("Courier New#b");
                if (isItalic) return new FontResolverInfo("Courier New#i");
                return new FontResolverInfo("Courier New#");
            }
            return new FontResolverInfo("Arial#");
        }
    }
} 