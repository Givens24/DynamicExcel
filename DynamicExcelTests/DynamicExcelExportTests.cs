using System;
using System.Collections.Generic;
using System.IO;
using DynamicExcel;
using DynamicExcelTests.MockDataModels;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DynamicExcelTests
{
    [TestClass]
    public class DynamicExcelExportTests
    {
        [TestMethod]
        [TestCategory("Integration")]
        [TestCategory("File IO")]
        public void Export_Data_Is_Exported_To_Excel_Successfully()
        {
            var customerMockData = BuildMockCustomerData();
            var testFilePath = Path.Combine(Directory.GetCurrentDirectory(), "test_excel_file.xlsx");
            var dynamicExcelExport = new DynamicExcelExport<CustomerOrderInfoModel>(customerMockData);
            dynamicExcelExport.Export(testFilePath);

            Assert.IsTrue(File.Exists(testFilePath));
            File.Delete(testFilePath);
        }

        private IEnumerable<CustomerOrderInfoModel> BuildMockCustomerData()
        {
            return new List<CustomerOrderInfoModel>
            {
                new CustomerOrderInfoModel
                {
                    Id = 654,
                    OrderNumber = 12348,
                    ProductId = Guid.NewGuid(),
                    CustomerName = "Jimmy Johns",
                    Address = "123 Happy Street",
                    ZipCode = 55124,
                    State = "MN"
                },
                new CustomerOrderInfoModel
                {
                    Id = 7845,
                    OrderNumber = 1897,
                    ProductId = Guid.NewGuid(),
                    CustomerName = "Target",
                    Address = "123 Sad Lane",
                    ZipCode = 55044,
                    State = "MN"
                },
                new CustomerOrderInfoModel
                {
                    Id = 9874,
                    OrderNumber = 8910,
                    ProductId = Guid.NewGuid(),
                    CustomerName = "Best Buy",
                    Address = "456 Pic- Me-Up Drive",
                    ZipCode = 55337,
                    State = "MN"
                }
            };
        }
    }
}
