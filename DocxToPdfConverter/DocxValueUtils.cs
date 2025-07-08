using NPOI.XWPF.UserModel;

namespace DocxToPdfConverter
{
    public static class DocxValueUtils
    {
        public static int ToIntSafe(object? value)
        {
            if (value == null) return 0;
            if (value is int i) return i;
            if (value is long l) return (int)l;
            if (value is ulong ul) return (int)ul;
            if (value is string s && !string.IsNullOrEmpty(s) && int.TryParse(s, out int result)) return result;
            return 0;
        }
        public static int GetSpacingAfterSafe(XWPFParagraph p)
        {
            if (p == null) return 0;
            return ToIntSafe(p.GetCTP()?.pPr?.spacing?.after);
        }
        public static int GetSpacingBeforeSafe(XWPFParagraph p)
        {
            if (p == null) return 0;
            return ToIntSafe(p.GetCTP()?.pPr?.spacing?.before);
        }
        public static int GetIndentationLeftSafe(XWPFParagraph p)
        {
            if (p == null) return 0;
            return ToIntSafe(p.GetCTP()?.pPr?.ind?.left);
        }
        public static int GetIndentationRightSafe(XWPFParagraph p)
        {
            if (p == null) return 0;
            return ToIntSafe(p.GetCTP()?.pPr?.ind?.right);
        }
    }
} 