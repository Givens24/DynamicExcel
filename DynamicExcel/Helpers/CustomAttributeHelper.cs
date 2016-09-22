using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using DynamicExcel.CustomAttributes;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel.Helpers
{
    public static class CustomAttributeHelper
    {
        internal static string GetColumnName(PropertyInfo propertyInfo)
        {
            var descriptionAttribute = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute)).FirstOrDefault();
            if (descriptionAttribute == null)
            {
                return propertyInfo.Name;
            }
            var description = descriptionAttribute as DescriptionAttribute;
            return description == null ? propertyInfo.Name : description.Description;
        }

        internal static int GetColumnWidth(PropertyInfo propertyInfo)
        {
            var columnAttribute = propertyInfo.GetCustomAttributes(typeof(ColumnWidth)).FirstOrDefault();
            var columnWidth = columnAttribute as ColumnWidth;
            return columnWidth?.Width ?? 20;
        }

        internal static int GetFontSize(PropertyInfo propertyInfo)
        {
            var fontAttribute = propertyInfo.GetCustomAttributes(typeof(FontSize)).FirstOrDefault();
            var fontSize = fontAttribute as FontSize;
            return fontSize?.Size ?? 11;
        }

        internal static Excel.XlRgbColor GetHeaderFont(PropertyInfo propertyInfo)
        {
            var headerFontColorAttribute = propertyInfo.GetCustomAttributes(typeof(HeaderFontColor)).FirstOrDefault();
            var headerFont = headerFontColorAttribute as HeaderFontColor;
            return headerFont?.FontColor ?? Excel.XlRgbColor.rgbBlack;
        }

        internal static Excel.XlRgbColor GetHeaderBackgroundColor(PropertyInfo propertyInfo)
        {
            var backgroundColorAttribute = propertyInfo.GetCustomAttributes(typeof(HeaderBackgroundColor)).FirstOrDefault();
            var backgroundColor = backgroundColorAttribute as HeaderBackgroundColor;
            return backgroundColor?.BackgroundColor ?? Excel.XlRgbColor.rgbWhite;
        }

        internal static Dictionary<Excel.XlHAlign, bool> GetColumnAlignment(PropertyInfo propertyInfo)
        {
            var alignmentOptions = new Dictionary<Excel.XlHAlign, bool>();
            var textAlignmentAttribute = propertyInfo.GetCustomAttributes(typeof(TextAlignment)).FirstOrDefault();
            var textAlignment = textAlignmentAttribute as TextAlignment;
            if (textAlignment == null)
            {
                alignmentOptions.Add(Excel.XlHAlign.xlHAlignLeft, false);
                return alignmentOptions;
            }

            alignmentOptions.Add(textAlignment.Alignment, textAlignment.EnitreColumn);

            return alignmentOptions;
        }
    }
}
