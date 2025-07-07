using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Internal;
using SixLabors.Fonts;

namespace XlsxToPdfConverter.Diy
{
    public class CustomFontResolver : IFontResolver
    {
        private static readonly string[] availableFontFiles;
        private static readonly Dictionary<string, FontFamilyModel> AvailableFonts =
            new Dictionary<string, FontFamilyModel>();

        static CustomFontResolver()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            availableFontFiles =
                Directory.GetFiles(Path.Combine(directory, "Fonts"), "*.ttf", SearchOption.AllDirectories);
            List<FontFileInfo> source = new List<FontFileInfo>();
            foreach (string fontFile in availableFontFiles)
            {
                try
                {
                    FontFileInfo fontFileInfo = FontFileInfo.Load(fontFile);
                    source.Add(fontFileInfo);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine((object)ex);
                }
            }

            foreach (IGrouping<string, FontFileInfo> fontList in source.GroupBy<FontFileInfo, string>(
                         (Func<FontFileInfo, string>)(info => info.FamilyName)))
            {
                try
                {
                    string key = fontList.Key;
                    FontFamilyModel fontFamilyModel = DeserializeFontFamily(key, (IEnumerable<FontFileInfo>)fontList);
                    AvailableFonts.Add(key.ToLower(), fontFamilyModel);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine((object)ex);
                }
            }
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (AvailableFonts.Count == 0)
            {
                throw new FileNotFoundException("No Fonts founded in bin/Fonts/ folder!");
            }

            // Явное сопоставление для Times New Roman
            if (familyName.Equals("Times New Roman", StringComparison.OrdinalIgnoreCase) ||
                familyName.Equals("Times", StringComparison.OrdinalIgnoreCase))
            {
                if (isBold && isItalic)
                    return new FontResolverInfo("timesbi.ttf");
                if (isBold)
                    return new FontResolverInfo("timesbd.ttf");
                if (isItalic)
                    return new FontResolverInfo("timesi.ttf");
                return new FontResolverInfo("times.ttf");
            }

            FontFamilyModel fontFamilyModel;
            if (AvailableFonts.TryGetValue(familyName.ToLower(), out fontFamilyModel))
            {
                if (isBold & isItalic)
                {
                    string path;
                    if (fontFamilyModel.FontFiles.TryGetValue(XFontStyleEx.BoldItalic, out path))
                    {
                        return new FontResolverInfo(Path.GetFileName(path));
                    }
                }
                else if (isBold)
                {
                    string path;
                    if (fontFamilyModel.FontFiles.TryGetValue(XFontStyleEx.Bold, out path))
                    {
                        return new FontResolverInfo(Path.GetFileName(path));
                    }
                }
                else
                {
                    string path;
                    if (isItalic && fontFamilyModel.FontFiles.TryGetValue(XFontStyleEx.Italic, out path))
                    {
                        return new FontResolverInfo(Path.GetFileName(path));
                    }
                }

                string path1;
                return fontFamilyModel.FontFiles.TryGetValue(XFontStyleEx.Regular, out path1)
                    ? new FontResolverInfo(Path.GetFileName(path1))
                    : new FontResolverInfo(Path.GetFileName(fontFamilyModel.FontFiles
                        .First<KeyValuePair<XFontStyleEx, string>>().Value));
            }

            return null;
        }

        public byte[] GetFont(string faceFileName)
        {
            using (MemoryStream destination = new MemoryStream())
            {
                string path = "";
                try
                {
                    path = availableFontFiles.First(x => 
                        string.Equals(Path.GetFileName(x), faceFileName, StringComparison.OrdinalIgnoreCase));
                    using (Stream stream = File.OpenRead(path))
                    {
                        stream.CopyTo(destination);
                        destination.Position = 0L;
                        return destination.ToArray();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine((object)ex);
                    throw new Exception("No Font File Found - " + faceFileName + " - " + path);
                }
            }
        }

        public string DefaultFontName => "Arial";

        private static FontFamilyModel DeserializeFontFamily(
            string fontFamilyName,
            IEnumerable<FontFileInfo> fontList)
        {
            FontFamilyModel fontFamilyModel = new FontFamilyModel() { Name = fontFamilyName };
            if (fontList.Count<FontFileInfo>() == 1)
            {
                fontFamilyModel.FontFiles.Add(XFontStyleEx.Regular, fontList.First<FontFileInfo>().Path);
            }
            else
            {
                foreach (FontFileInfo font in fontList)
                {
                    XFontStyleEx key = font.GuessFontStyle();
                    if (!fontFamilyModel.FontFiles.ContainsKey(key))
                    {
                        fontFamilyModel.FontFiles.Add(key, font.Path);
                    }
                }
            }

            return fontFamilyModel;
        }

        private readonly struct FontFileInfo
        {
            private FontFileInfo(string path, FontDescription fontDescription)
            {
                this.Path = path;
                this.FontDescription = fontDescription;
            }

            public string Path { get; }

            public FontDescription FontDescription { get; }

            public string FamilyName => FontDescription.FontFamilyInvariantCulture;

            public XFontStyleEx GuessFontStyle()
            {
                switch (FontDescription.Style)
                {
                    case FontStyle.Bold:
                        return XFontStyleEx.Bold;
                    case FontStyle.Italic:
                        return XFontStyleEx.Italic;
                    case FontStyle.BoldItalic:
                        return XFontStyleEx.BoldItalic;
                    default:
                        return XFontStyleEx.Regular;
                }
            }

            internal static FontFileInfo Load(string path)
            {
                FontDescription fontDescription = FontDescription.LoadDescription(path);
                return new FontFileInfo(path, fontDescription);
            }
        }
    }
}