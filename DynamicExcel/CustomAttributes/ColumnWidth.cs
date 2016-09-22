using System;

namespace DynamicExcel.CustomAttributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnWidth : Attribute
    {
        private readonly int _columnWidth;
        public ColumnWidth(int columnWidth)
        {
            _columnWidth = columnWidth;
        }

        public int Width => _columnWidth;
    }
}
