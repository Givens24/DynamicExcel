using System;
using System.ComponentModel;
using DynamicExcel.CustomAttributes;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcelTests.MockDataModels
{
    public class CustomerOrderInfoModel
    {
        [HeaderFontColor(Excel.XlRgbColor.rgbWhite)]
        [HeaderBackgroundColor(Excel.XlRgbColor.rgbNavyBlue)]
        [Description("Customer Id")]
        public int Id { get; set; }
        [Description("Order Number")]
        public int OrderNumber { get; set; }
        [FontSize(20)]
        [Description("Customer Name")]
        public string CustomerName { get; set; }
        [ColumnWidth(35)]
        public string Address { get; set; }
        [Description("Zip Code")]
        public long ZipCode { get; set; }
        public string State { get; set; }
        [Description("Product Id")]
        [TextAlignment(Excel.XlHAlign.xlHAlignCenter, true)]
        public Guid ProductId { get; set; }
    }
}
