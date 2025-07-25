#nullable enable

namespace DocxToPdfConverter
{
    public static class DocxValueUtils
    {
        internal static int ToIntSafe(object? value)
        {
            return value switch
            {
                null => 0,
                int i => i,
                long l => (int)l,
                ulong ul => (int)ul,
                string s when !string.IsNullOrEmpty(s) && int.TryParse(s, out var result) => result,
                _ => 0
            };
        }
    }
} 