using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace ExcelToHtml.Core
{
    internal sealed class CellStyleReader
    {
        private const string ItalicClassName = "e2h-italic";
        private const string UnderLineClassName = "e2h-underline";
        private const string BoldClassName = "e2h-bold";

        private const string WrapTextClassName = "e2h-wrap";

        private const string DefaultFontColor = "#000000";
        private const string DefaultFillColor = "#ffffff";
        private const float DefaultFontSize = 12;

        private readonly Dictionary<string, string> _colors = new();
        private readonly Dictionary<string, string> _fillColors = new();
        private readonly HashSet<float> _fontSize = new();

        internal CellStyle Convert(ExcelRange cell)
        {
            var style = new CellStyle();

            var fontColor = ExcelColorToColor(cell.Style.Font.Color);
            //默认字体颜色为黑色
            if (fontColor != DefaultFontColor)
            {
                if (!_colors.ContainsKey(fontColor))
                {
                    var className = $"e2h-c-{_colors.Count}";
                    _colors.Add(fontColor, className);
                    style.ClassNames.Add(className);
                }
                else
                {
                    style.ClassNames.Add(_colors[fontColor]);
                }
            }

            if (cell.Style.Font.Bold && !style.ClassNames.Contains(BoldClassName))
                style.ClassNames.Add(BoldClassName);
            if (cell.Style.Font.Italic && !style.ClassNames.Contains(ItalicClassName))
                style.ClassNames.Add(ItalicClassName);
            if (cell.Style.Font.UnderLine && !style.ClassNames.Contains(UnderLineClassName))
                style.ClassNames.Add(UnderLineClassName);

            if (cell.Style.Font.Size != DefaultFontSize && cell.Style.Font.Size>DefaultFontSize)
            {
                if (!_fontSize.Contains(cell.Style.Font.Size))
                    _fontSize.Add(cell.Style.Font.Size);
                style.ClassNames.Add($"fs-{cell.Style.Font.Size}");
            }


            if (cell.Style.WrapText)
                style.ClassNames.Add(WrapTextClassName);
            if (!string.IsNullOrWhiteSpace(cell.Style.Fill.BackgroundColor.Rgb))
            {
                var fillColor = ExcelColorToColor(cell.Style.Fill.BackgroundColor);
                if (fillColor !=DefaultFillColor)
                {
                    if (!_fillColors.ContainsKey(fillColor))
                    {
                        var className = $"e2h-bg-{_fillColors.Count}";
                        _fillColors.Add(fillColor, className);
                        style.ClassNames.Add(className);
                    }
                    else
                    {
                        style.ClassNames.Add(_fillColors[fillColor]);
                    }
                }
            }

            switch (cell.Style.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Center:
                    style.Align = "center";
                    break;
                case ExcelHorizontalAlignment.Right:
                    style.Align = "right";
                    break;
            }

            switch (cell.Style.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Bottom:
                    style.Valign = "bottom";
                    break;
                case ExcelVerticalAlignment.Top:
                    style.Valign = "top";
                    break;

            }

            return style;
        }

        public IReadOnlyDictionary<string, string> GetAllFontColors()
        {
            return _colors;
        }

        public IReadOnlyDictionary<string, string> GetAllFillColors()
        {
            return _fillColors;
        }

        public HashSet<float> GetAllFontSize()
        {
            return _fontSize;
        }

        public static string ColorToHexValue(Color color)
        {
            return "#" + color.R.ToString("X2") +
                   color.G.ToString("X2") +
                   color.B.ToString("X2");
        }

        public static string ExcelColorToColor(ExcelColor color)
        {
            if (!string.IsNullOrWhiteSpace(color.Rgb))
                return $"#{color.Rgb.Substring(2).ToLower()}";
            var value = color.LookupColor();
            return $"#{value.Substring(3).ToLower()}";
        }
    }

}
