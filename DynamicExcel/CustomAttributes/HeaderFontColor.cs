using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel.CustomAttributes
{
    public class HeaderFontColor : Attribute
    {
        private readonly Excel.XlRgbColor _fontColor;

        public HeaderFontColor(Excel.XlRgbColor backgroundColor)
        {
            _fontColor = backgroundColor;
        }

        public Excel.XlRgbColor FontColor => _fontColor;
    }
}
