using System;

namespace XlsxToPdfConverter.Diy
{
    public interface ILogger
    {
        void Log(string message);
        void Error(Exception ex, string message);
        void Warn(string message);
    }
} 