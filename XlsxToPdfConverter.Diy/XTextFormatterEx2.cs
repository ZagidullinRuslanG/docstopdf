//http://developer.th-soft.com/developer/2015/09/21/xtextformatter-revisited-xtextformatterex2-for-pdfsharp-1-50-beta-2/

using System;
using System.Collections.Generic;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Represents a very simple text formatter.
    /// If this class does not satisfy your needs on formatting paragraphs I recommend to take a look
    /// at MigraDoc Foundation. Alternatively you should copy this class in your own source code and modify it.
    /// </summary>
    public class XTextFormatterEx2
    {
        public enum SpacingMode
        {
            /// <summary>
            /// With Relative, the value of Spacing will be added to the default line space.
            /// With 0 you get the default behaviour.
            /// With 5 the line spacing will be 5 points larger than the default spacing.
            /// </summary>
            Relative,

            /// <summary>
            /// With Absolute you set the absolute line spacing.
            /// With 0 all the text will be written at the same line.
            /// </summary>
            Absolute,

            /// <summary>
            /// With Percentage, you can specify larger or smaller line spacing.
            /// With 100 you get the default behaviour.
            /// With 200 you get double line spacing.
            /// With 90 you get 90% of the default line spacing.
            /// </summary>
            Percentage,
        }

        public struct LayoutOptions
        {
            public float Spacing;

            public SpacingMode SpacingMode;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XTextFormatter"/> class.
        /// </summary>
        public XTextFormatterEx2(XGraphics gfx)
            : this(gfx, new LayoutOptions { SpacingMode = SpacingMode.Relative, Spacing = 0 })
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XTextFormatter"/> class.
        /// </summary>
        public XTextFormatterEx2(XGraphics gfx, LayoutOptions options)
        {
            if (gfx == null)
            {
                throw new ArgumentNullException("gfx");
            }

            this.gfx = gfx;
            layoutOptions = options;
        }

        private readonly XGraphics gfx;
        private readonly LayoutOptions layoutOptions;

        private bool preparedText;

        /// <summary>
        /// Gets or sets the bounding box of the layout.
        /// </summary>
        public XRect LayoutRectangle
        {
            get { return layoutRectangle; }
            set { layoutRectangle = value; }
        }

        private XRect layoutRectangle;

        /// <summary>
        /// Gets or sets the alignment of the text.
        /// </summary>
        public XParagraphAlignment Alignment
        {
            get { return alignment; }
            set { alignment = value; }
        }

        private XParagraphAlignment alignment = XParagraphAlignment.Left;

        /// <summary>
        /// Prepares a given text for drawing, performs the layout, returns the index of the last fitting char and the needed height.
        /// </summary>
        /// <param name="text">The text to be drawn.</param>
        /// <param name="font">The font to be used.</param>
        /// <param name="brush">The brush to be used.</param>
        /// <param name="layoutRectangle">The layout rectangle. Set the correct width.
        /// Either set the available height to find how many chars will fit.
        /// Or set height to double.MaxValue to find which height will be needed to draw the whole text.</param>
        /// <param name="lastFittingChar">Index of the last fitting character. Can be -1 if the character was not determined. Will be -1 if the whole text can be drawn.</param>
        /// <param name="neededHeight">The needed height - either for the complete text or the used height of the given rect.</param>
        /// <exception cref="ArgumentNullException">Text or font cannot be null.</exception>
        public void PrepareDrawString(string text, XFont font, XBrush brush, XRect layoutRectangle, out int lastFittingChar, out double neededHeight)
        {
            LayoutRectangle = layoutRectangle;
            blocks.Clear();

            AppendPreparedDrawString(text, font, brush, out lastFittingChar, out neededHeight);
        }

        /// <summary>
        /// Prepares a given text for drawing, performs the layout, returns the index of the last fitting char and the needed height.
        /// </summary>
        /// <param name="text">The text to be drawn.</param>
        /// <param name="font">The font to be used.</param>
        /// Either set the available height to find how many chars will fit.
        /// Or set height to double.MaxValue to find which height will be needed to draw the whole text.</param>
        /// <param name="lastFittingChar">Index of the last fitting character. Can be -1 if the character was not determined. Will be -1 if the whole text can be drawn.</param>
        /// <param name="neededHeight">The needed height - either for the complete text or the used height of the given rect.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void AppendPreparedDrawString(string text, XFont font, XBrush brush, out int lastFittingChar, out double neededHeight)
        {
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }

            if (font == null)
            {
                throw new ArgumentNullException("font");
            }

            if (layoutRectangle == default)
            {
                throw new ArgumentException("LayoutRectangle must be set first.");
            }

            lastFittingChar = -1;
            neededHeight = double.MinValue;

            if (text.Length == 0)
            {
                return;
            }

            CreateBlocks(text, font, brush);

            CreateLayout();

            preparedText = true;

            int count = blocks.Count;
            for (int idx = 0; idx < count; idx++)
            {
                Block block = blocks[idx];
                if (block.Stop)
                {
                    // We have a Stop block, so only part of the text will fit. We return the index of the last fitting char (and the height of the block, if available).
                    lastFittingChar = 0;
                    int idx2 = idx - 1;
                    while (idx2 >= 0)
                    {
                        Block block2 = blocks[idx2];
                        if (block2.EndIndex >= 0)
                        {
                            lastFittingChar = block2.EndIndex;
                            neededHeight = 0;
                            for (int idx3 = 0; idx3 <= idx2; idx3++)
                            {
                                Block block3 = blocks[idx3];
                                neededHeight = Math.Max(neededHeight, (block3.Style?.CyDescent ?? 0) + block3.Location.Y); // Test this!!!!!
                            }

                            return;
                        }
                        --idx2;
                    }
                    return;
                }
                if (block.Type == BlockType.LineBreak)
                {
                    continue;
                }

                //gfx.DrawString(block.Text, font, brush, dx + block.Location.x, dy + block.Location.y);
                neededHeight = Math.Max(neededHeight, (block.Style?.CyDescent ?? 0) + block.Location.Y); // Test this!!!!! Performance optimization?
            }
        }

        /// <summary>
        /// Draws the text that was previously prepared by calling PrepareDrawString or by passing a text to DrawString.
        /// </summary>
        /// <exception cref="ArgumentException">PrepareDrawString must be called first.</exception>
        public void DrawString()
        {
            // TODO: Do we need "XStringFormat format" at PrepareDrawString or at DrawString? Not yet used anyway, but probably already needed at PrepareDrawString.
            if (!preparedText)
            {
                throw new ArgumentException("PrepareDrawString must be called first.");
            }

            if (blocks.Count == 0)
            {
                return;
            }

            double dx = layoutRectangle.Location.X;
            double dy = layoutRectangle.Location.Y;
            int count = blocks.Count;
            for (int idx = 0; idx < count; idx++)
            {
                Block block = blocks[idx];
                if (block.Stop)
                {
                    break;
                }

                if (block.Type == BlockType.LineBreak)
                {
                    continue;
                }

                gfx.DrawString(block.Text, block.Style.Font, block.Style.Brush, dx + block.Location.X, dy + block.Location.Y);
            }
        }

        /// <summary>
        /// Draws the text.
        /// </summary>
        /// <param name="text">The text to be drawn.</param>
        /// <param name="font">The font.</param>
        /// <param name="brush">The text brush.</param>
        /// <param name="layoutRectangle">The layout rectangle.</param>
        public void DrawString(string text, XFont font, XBrush brush, XRect layoutRectangle)
        {
            DrawString(text, font, brush, layoutRectangle, XStringFormats.TopLeft);
        }

        /// <summary>
        /// Draws the text.
        /// </summary>
        /// <param name="text">The text to be drawn.</param>
        /// <param name="font">The font.</param>
        /// <param name="brush">The text brush.</param>
        /// <param name="layoutRectangle">The layout rectangle.</param>
        /// <param name="format">The format. Must be <c>XStringFormat.TopLeft</c>.</param>
        public void DrawString(string text, XFont font, XBrush brush, XRect layoutRectangle, XStringFormat format)
        {
            int dummy1;
            double dummy2;
            PrepareDrawString(text, font, brush, layoutRectangle, out dummy1, out dummy2);

            DrawString();
        }

        private void CreateBlocks(string text, XFont font, XBrush brush)
        {
            int length = text.Length;
            bool inNonWhiteSpace = false;
            int startIndex = 0, blockLength = 0;
            var style = Style.Create(gfx, font, brush, layoutOptions);
            for (int idx = 0; idx < length; idx++)
            {
                char ch = text[idx];

                // Treat CR and CRLF as LF
                if (ch == Chars.CR)
                {
                    if (idx < length - 1 && text[idx + 1] == Chars.LF)
                    {
                        idx++;
                    }

                    ch = Chars.LF;
                }
                if (ch == Chars.LF)
                {
                    if (blockLength != 0)
                    {
                        string token = text.Substring(startIndex, blockLength);
                        blocks.Add(new Block(token, BlockType.Text,
                          gfx.MeasureString(token, font).Width,
                          startIndex, startIndex + blockLength - 1, style));
                    }
                    startIndex = idx + 1;
                    blockLength = 0;
                    blocks.Add(new Block(BlockType.LineBreak));
                }
                else if (char.IsWhiteSpace(ch))
                {
                    if (inNonWhiteSpace)
                    {
                        string token = text.Substring(startIndex, blockLength);
                        blocks.Add(new Block(token, BlockType.Text,
                          gfx.MeasureString(token, font).Width,
                          startIndex, startIndex + blockLength - 1, style));
                        startIndex = idx + 1;
                        blockLength = 0;
                    }
                    else
                    {
                        blockLength++;
                    }
                }
                else
                {
                    inNonWhiteSpace = true;
                    blockLength++;
                }
            }
            if (blockLength != 0)
            {
                string token = text.Substring(startIndex, blockLength);
                blocks.Add(new Block(token, BlockType.Text,
                                gfx.MeasureString(token, font).Width,
                                startIndex, startIndex + blockLength - 1, style));
            }
        }

        private void CreateLayout()
        {
            double rectWidth = layoutRectangle.Width;
            double rectHeight = layoutRectangle.Height;
            int firstIndex = 0;
            double x = 0, y = 0;
            int count = blocks.Count;
            for (int idx = 0; idx < count; idx++)
            {
                Block block = blocks[idx];
                if (block.Type == BlockType.LineBreak)
                {
                    if (Alignment == XParagraphAlignment.Justify)
                    {
                        blocks[firstIndex].Alignment = XParagraphAlignment.Left;
                    }

                    AlignLine(firstIndex, idx - 1, rectWidth);
                    x = 0;
                    y += GetLineSpace(firstIndex, idx - 1);
                    var cyDescent = GetCyDescent(firstIndex, idx - 1);
                    firstIndex = idx + 1;
                    if (y + cyDescent > rectHeight)
                    {
                        block.Stop = true;
                        break;
                    }
                }
                else
                {
                    double width = block.Width; //!!!modTHHO 19.11.09 don't add this.spaceWidth here
                    if ((x + width <= rectWidth || x == 0) && block.Type != BlockType.LineBreak)
                    {
                        block.Location = new XPoint(x, y);
                        x += width + block.Style.SpaceWidth; //!!!modTHHO 19.11.09 add this.spaceWidth here
                    }
                    else
                    {
                        AlignLine(firstIndex, idx - 1, rectWidth);
                        y += GetLineSpace(firstIndex, idx - 1);
                        var cyDescent = GetCyDescent(firstIndex, idx - 1);
                        firstIndex = idx;
                        if (y + cyDescent > rectHeight)
                        {
                            block.Stop = true;
                            break;
                        }
                        block.Location = new XPoint(0, y);
                        x = width + block.Style.SpaceWidth; //!!!modTHHO 19.11.09 add this.spaceWidth here
                    }
                }
            }
            if (firstIndex < count && Alignment != XParagraphAlignment.Justify)
            {
                AlignLine(firstIndex, count - 1, rectWidth);
            }
        }

        private double GetLineSpace(int firstIndex, int lastIndex)
        {
            double lineSpace = 0;
            for (int idx = firstIndex; idx <= lastIndex; idx++)
            {
                var effectiveLineSpace = blocks[idx].Style?.EffectiveLineSpace;
                if (effectiveLineSpace.HasValue)
                {
                    lineSpace = Math.Max(lineSpace, effectiveLineSpace.Value);
                }
            }

            return lineSpace;
        }

        private double GetCyDescent(int firstIndex, int lastIndex)
        {
            double cyDescent = 0;
            for (int idx = firstIndex; idx <= lastIndex; idx++)
            {
                var descent = blocks[idx].Style?.CyDescent;
                if (descent.HasValue)
                {
                    cyDescent = Math.Max(cyDescent, descent.Value);
                }
            }

            return cyDescent;
        }

        /// <summary>
        /// Align center, right or justify.
        /// </summary>
        private void AlignLine(int firstIndex, int lastIndex, double layoutWidth)
        {
            int count = lastIndex - firstIndex + 1;
            if (count == 0)
            {
                return;
            }

            double totalWidth = -blocks[lastIndex].Style?.SpaceWidth ?? 0;
            double cyAccent = 0;
            for (int idx = firstIndex; idx <= lastIndex; idx++)
            {
                totalWidth += blocks[idx].Width + blocks[idx].Style?.SpaceWidth ?? 0;
                cyAccent = Math.Max(cyAccent, blocks[idx].Style?.CyAscent ?? 0);
            }

            double dx = Math.Max(layoutWidth - totalWidth, 0);
            //Debug.Assert(dx >= 0);
            if (alignment != XParagraphAlignment.Justify)
            {
                if (alignment == XParagraphAlignment.Center)
                {
                    dx /= 2;
                }
                else if (alignment == XParagraphAlignment.Left)
                {
                    dx = 0;
                }

                for (int idx = firstIndex; idx <= lastIndex; idx++)
                {
                    Block block = blocks[idx];
                    block.Location += new XSize(dx, cyAccent);
                }
            }
            // case: justify
            else if (count > 1)
            {
                dx /= count - 1;
                for (int idx = firstIndex + 1, i = 1; idx <= lastIndex; idx++, i++)
                {
                    Block block = blocks[idx];
                    block.Location += new XSize(dx * i, 0);
                }
            }
        }

        private readonly List<Block> blocks = new List<Block>();

        private enum BlockType
        {
            Text,
            Space,
            Hyphen,
            LineBreak,
        }

        private class Style
        {
            public readonly XFont Font;
            public readonly XBrush Brush;
            public readonly double LineSpace;
            public readonly double EffectiveLineSpace;
            public readonly double CyAscent;
            public readonly double CyDescent;
            public readonly double SpaceWidth;

            private Style(XFont font, XBrush brush, double lineSpace, double effectiveLineSpace, double cyAscent, double cyDescent, double spaceWidth)
            {
                Font = font;
                Brush = brush;
                LineSpace = lineSpace;
                EffectiveLineSpace = effectiveLineSpace;
                CyAscent = cyAscent;
                CyDescent = cyDescent;
                SpaceWidth = spaceWidth;
            }

            public static Style Create(XGraphics gfx, XFont font, XBrush brush, LayoutOptions layoutOptions)
            {
                var lineSpace = font.GetHeight();
                var effectiveLineSpace = CalculateLineSpace(lineSpace, layoutOptions);
                var cyAscent = lineSpace * font.CellAscent / font.CellSpace;
                var cyDescent = lineSpace * font.CellDescent / font.CellSpace;

                // HACK in XTextFormatter
                var spaceWidth = gfx.MeasureString("x x", font).Width;
                spaceWidth -= gfx.MeasureString("xx", font).Width;

                return new Style(font, brush, lineSpace, effectiveLineSpace, cyAscent, cyDescent, spaceWidth);
            }

            private static double CalculateLineSpace(double lineSpace, LayoutOptions layoutOptions)
            {
                switch (layoutOptions.SpacingMode)
                {
                    case SpacingMode.Absolute:
                        return layoutOptions.Spacing;
                    case SpacingMode.Relative:
                        return lineSpace + layoutOptions.Spacing;
                    case SpacingMode.Percentage:
                        return lineSpace * layoutOptions.Spacing / 100;
                }

                return lineSpace;
            }
        }

        /// <summary>
        /// Represents a single word.
        /// </summary>
        private class Block
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="Block"/> class.
            /// </summary>
            /// <param name="text">The text of the block.</param>
            /// <param name="type">The type of the block.</param>
            /// <param name="width">The width of the text.</param>
            public Block(string text, BlockType type, double width, int startIndex, int endIndex, Style style)
            {
                Text = text;
                Type = type;
                Width = width;
                StartIndex = startIndex;
                EndIndex = endIndex;
                Style = style;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Block"/> class.
            /// </summary>
            /// <param name="type">The type.</param>
            public Block(BlockType type)
            {
                Type = type;
            }

            /// <summary>
            /// The text represented by this block.
            /// </summary>
            public readonly string Text;

            public readonly int StartIndex = -1;
            public readonly int EndIndex = -1;
            public readonly Style Style;

            /// <summary>
            /// The type of the block.
            /// </summary>
            public readonly BlockType Type;

            /// <summary>
            /// The width of the text.
            /// </summary>
            public readonly double Width;

            /// <summary>
            /// The location relative to the upper left corner of the layout rectangle.
            /// </summary>
            public XPoint Location;

            /// <summary>
            /// The alignment of this line.
            /// </summary>
            public XParagraphAlignment Alignment;

            /// <summary>
            /// A flag indicating that this is the last block that fits in the layout rectangle.
            /// </summary>
            public bool Stop;
        }
        // TODO:
        // - more XStringFormat variations
        // - calculate bounding box
        // - left and right indent
        // - first line indent
        // - margins and paddings
        // - background color
        // - text background color
        // - border style
        // - hyphens, soft hyphens, hyphenation
        // - kerning
        // - change font, size, text color etc.
        // - line spacing
        // - underline and strike-out variation
        // - super- and sub-script
        // - ...
    }
}
