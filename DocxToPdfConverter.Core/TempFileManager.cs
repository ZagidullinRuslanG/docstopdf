namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Реализация менеджера временных файлов
    /// </summary>
    public class TempFileManager : ITempFileManager
    {
        private readonly List<string> _tempFiles = new();
        private readonly ILogger _logger;
        private bool _disposed = false;

        public TempFileManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string CreateTempFile(string extension = ".tmp")
        {
            var tempPath = Path.GetTempFileName();
            var newPath = Path.ChangeExtension(tempPath, extension);
            
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
            
            RegisterTempFile(newPath);
            return newPath;
        }

        public void RegisterTempFile(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath) && !_tempFiles.Contains(filePath))
            {
                _tempFiles.Add(filePath);
            }
        }

        public void Cleanup()
        {
            foreach (var tempFile in _tempFiles.ToList())
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                        _logger.LogInfo($"Удален временный файл: {tempFile}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось удалить временный файл {tempFile}: {ex.Message}");
                }
            }
            _tempFiles.Clear();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Cleanup();
                _disposed = true;
            }
        }
    }
}
