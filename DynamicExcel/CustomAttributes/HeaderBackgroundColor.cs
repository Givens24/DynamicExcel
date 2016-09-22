using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel.CustomAttributes
{
    public class HeaderBackgroundColor : Attribute
    {
        private readonly Excel.XlRgbColor _backgroundColor;

        public HeaderBackgroundColor(Excel.XlRgbColor backgroundColor)
        {
            _backgroundColor = backgroundColor;
        }

        public Excel.XlRgbColor BackgroundColor => _backgroundColor;
    }
}
