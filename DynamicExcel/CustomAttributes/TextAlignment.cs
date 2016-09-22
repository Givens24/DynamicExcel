using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel.CustomAttributes
{
    public class TextAlignment: Attribute
    {
        private readonly Excel.XlHAlign _alignment;
        private readonly bool _entireColumn;
        public TextAlignment(Excel.XlHAlign alignment, bool entireColumn)
        {
            _entireColumn = entireColumn;
            _alignment = alignment;
        }

        public Excel.XlHAlign Alignment => _alignment;
        public bool EnitreColumn => _entireColumn;
    }
}
