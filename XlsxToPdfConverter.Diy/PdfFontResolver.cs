using PdfSharp.Fonts;
using System;
using System.IO;
using System.Reflection;

namespace XlsxToPdfConverter.Diy
{
    public class PdfFontResolver : IFontResolver
    {
        private static readonly string FontsFolder = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "",
            "Fonts");

        public byte[] GetFont(string faceName)
        {
            string fontFile = faceName switch
            {
                "Arial#" => "arial.ttf",
                "Arial#b" => "arialbd.ttf",
                "Arial#i" => "ariali.ttf",
                "Arial#bi" => "arialbi.ttf",
                _ => throw new NotImplementedException($"Font face '{faceName}' not implemented.")
            };

            string path = Path.Combine(FontsFolder, fontFile);
            return File.ReadAllBytes(path);
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName.Equals("Arial", StringComparison.OrdinalIgnoreCase))
            {
                if (isBold && isItalic)
                    return new FontResolverInfo("Arial#bi");
                if (isBold)
                    return new FontResolverInfo("Arial#b");
                if (isItalic)
                    return new FontResolverInfo("Arial#i");
                return new FontResolverInfo("Arial#");
            }
            // Можно добавить поддержку других шрифтов, если потребуется
            return null;
        }
    }
} 