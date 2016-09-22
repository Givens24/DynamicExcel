using System;

namespace DynamicExcel.CustomAttributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FontSize : Attribute
    {
        private readonly int _fontSize;

        public FontSize(int fontSize)
        {
            _fontSize = fontSize;
        }

        public int Size => _fontSize;
    }
}
