using System.Collections.Generic;
using System.Linq;
using DynamicExcel.Helpers;
using DynamicExcel.Models;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel
{
    public abstract class DynamicExcelBase
    {
        public readonly Excel._Application Application;
        protected DynamicExcelBase()
        {
            Application = new Excel.Application();
        }

        protected void StartApplication()
        {
            Application.Run();
        }

        protected void QuitApplication()
        {
            Application.Quit();
        }

        protected Excel.Workbook CreateNewWorkbook()
        {
            var workBook = Application.Workbooks.Add();
            return workBook;
        }

        protected IEnumerable<DynamicExcelColumn> GetExcelColumns<T>() where T : class
        {
            var columns = new List<DynamicExcelColumn>();
            var dataTypeProperties = typeof(T).GetProperties();
            dataTypeProperties.ToList().ForEach(x =>
            {
                columns.Add(new DynamicExcelColumn
                {
                    Name = CustomAttributeHelper.GetColumnName(x),
                    Width = CustomAttributeHelper.GetColumnWidth(x),
                    FontSize = CustomAttributeHelper.GetFontSize(x),
                    HeaderBackgroundColor = CustomAttributeHelper.GetHeaderBackgroundColor(x),
                    HeaderFontColor = CustomAttributeHelper.GetHeaderFont(x),
                    Alignment = CustomAttributeHelper.GetColumnAlignment(x)
                });
            });

            return columns;
        }

        protected void FormatColumnHeader(DynamicExcelColumn column, Excel.Range headerColumnCell)
        {
            headerColumnCell.Interior.Color = column.HeaderBackgroundColor;
            headerColumnCell.Font.Color = column.HeaderFontColor;
            headerColumnCell.ColumnWidth = column.Width;
            headerColumnCell.Font.Size = column.FontSize;
            headerColumnCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            headerColumnCell.Borders.Color = ColorTranslator.ToOle(Color.SlateGray);
            SetTextAlignment(column, headerColumnCell);
        }

        private void SetTextAlignment(DynamicExcelColumn column, Excel.Range headerColumnCell)
        {
            var headerAlignment = column.Alignment.FirstOrDefault();
            if (headerAlignment.Value)
            {
                headerColumnCell.EntireColumn.HorizontalAlignment = headerAlignment.Key;
            }
            else
            {
                headerColumnCell.HorizontalAlignment = headerAlignment.Key;
            }
        }
    }
}
