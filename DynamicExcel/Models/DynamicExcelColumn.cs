using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel.Models
{
    public class DynamicExcelColumn
    {
        internal string Name { get; set; }
        internal int Width { get; set; }
        internal int FontSize { get; set; }
        internal Dictionary<Excel.XlHAlign , bool> Alignment { get; set; }
        internal Excel.XlRgbColor HeaderFontColor { get; set; }
        internal Excel.XlRgbColor HeaderBackgroundColor { get; set; }
    }
}
