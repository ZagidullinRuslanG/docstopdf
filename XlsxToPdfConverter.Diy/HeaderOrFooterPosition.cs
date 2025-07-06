namespace XlsxToPdfConverter.Diy
{
    public enum HeaderOrFooterPosition
    {
        /// <summary>
        /// Слева.
        /// </summary>
        Left = 0,
        
        /// <summary>
        /// По центру.
        /// </summary>
        Center = 1,
        
        /// <summary>
        /// Справа.
        /// </summary>
        Right = 2,

        /// <summary>
        /// Сверху.
        /// </summary>
        Top = 4,
        
        /// <summary>
        /// Снизу.
        /// </summary>
        Bottom = 8,

        /// <summary>
        /// Сверху слева.
        /// </summary>
        TopLeft = Top | Left,
        
        /// <summary>
        /// Сверху по центру.
        /// </summary>
        TopCenter = Top | Center,
        
        /// <summary>
        /// Сверху справа.
        /// </summary>
        TopRight = Top | Right,
        
        /// <summary>
        /// Снизу слева.
        /// </summary>
        BottomLeft = Bottom | Left,
        
        /// <summary>
        /// Снизу по центру.
        /// </summary>
        BottomCenter = Bottom | Center,
        
        /// <summary>
        /// Снизу справа.
        /// </summary>
        BottomRight = Bottom | Right,
    }
}