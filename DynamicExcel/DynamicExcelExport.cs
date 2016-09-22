using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicExcel
{
    public class DynamicExcelExport<T> : DynamicExcelBase where T : class
    {
        private readonly IEnumerable<T> _dataToExport;
        private readonly Excel.Workbook _workBook;
        private readonly List<PropertyInfo> _properties;
        public DynamicExcelExport(IEnumerable<T> dataToExport)
        {
            _dataToExport = dataToExport;
            _workBook = CreateNewWorkbook();
            _properties = typeof(T).GetProperties().ToList();
            if (ClassHasNestedCollections())
            {
                throw new InvalidOperationException("Class cannot have nested collections.");
            }
        }

        private bool ClassHasNestedCollections()
        {
            //var hasNestedCollections = false;
            //foreach (var property in _properties)
            //{
            //    if (property.PropertyType == typeof(string))
            //    {
            //        continue;
            //    }

            //    if (typeof(IEnumerable).IsAssignableFrom(property.PropertyType) || 
            //        typeof(IEnumerable<>).IsAssignableFrom(property.PropertyType))
            //    {
            //        hasNestedCollections = true;
            //    }
            //}

            //return hasNestedCollections;

            return _properties.ToList().Any(x => x.PropertyType != typeof(string) &&
                                                 typeof(IEnumerable).IsAssignableFrom(x.PropertyType) ||
                                                 typeof(IEnumerable<>).IsAssignableFrom(x.PropertyType));
        }

        public void Export(string filePath, string worksheetName = "", int workSheetIndex = 1)
        {
            Excel.Worksheet workSheet = _workBook.Sheets[workSheetIndex];
            if (!string.IsNullOrEmpty(worksheetName))
            {
                workSheet.Name = worksheetName;
            }

            ExportColumnHeaders(workSheet);
            ExportDataToSpreadSheet(workSheet);
            SaveWorkbook(filePath);
        }

        private void SaveWorkbook(string filePath)
        {
            if (File.Exists(filePath))
            {
                throw new InvalidOperationException($"An error occurred while exporting the data to Excel. Error: The file {filePath} already exists.");
            }
            _workBook.SaveAs(filePath);
            _workBook.Close();
        }

        private void ExportColumnHeaders(Excel.Worksheet workSheet)
        {
            var columns = GetExcelColumns<T>();
            var columnIndex = 1;
            columns.ToList().ForEach(x =>
            {
                var cell = (Excel.Range) workSheet.Cells[1, columnIndex];
                FormatColumnHeader(x, cell);
                cell.Value2 = x.Name;
                columnIndex++;
            });
        }

        private void ExportDataToSpreadSheet(Excel.Worksheet workSheet)
        {
            var rowIndex = 2;
            _dataToExport.ToList().ForEach(x =>
            {
                SetCellValues(workSheet, rowIndex, x);
                rowIndex++;
            });
        }

        private void SetCellValues(Excel.Worksheet workSheet, int rowIndex, T x)
        {
            var columnIndex = 1;
            _properties.ForEach(property =>
            {
                var cell = (Excel.Range) workSheet.Cells[rowIndex, columnIndex];
                SetCellValue(property, cell, x);
                columnIndex++;
            });
        }

        private void SetCellValue(PropertyInfo property, Excel.Range cell, T data)
        {
            var value = property.GetValue(data, null);
            cell.Value2 = value is Guid ? value.ToString() : value;
        }
    }
}