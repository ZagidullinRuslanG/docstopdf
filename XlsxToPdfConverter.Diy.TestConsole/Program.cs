using System;
using System.IO;

namespace XlsxToPdfConverter.Diy.TestConsole
{
    internal class Program
    {
        // note: Консольное приложение для профайлера.

        private static void Main(string[] args)
        {
            string dir = args[0];
            IXlsxToPdfConverter converter = new XlsxToPdfDiyConverter();
            string jobId = Guid.NewGuid().ToString("N");
            foreach (string file in Directory.GetFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                converter.Convert(file, file + $".{jobId}.pdf");
            }
        }
    }
}
