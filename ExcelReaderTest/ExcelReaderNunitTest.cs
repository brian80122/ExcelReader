using ExcelReaderLibrary;
using ExcelReaderTest.Models;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelReaderTest
{
    [TestFixture()]
    public class ExcelReaderNunitTest
    {
        private IExcelReader<Product> _excelReader { get; set; }
        private readonly string _baseFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelSamples");
        public ExcelReaderNunitTest()
        {
            _excelReader = new ExcelReader<Product>();
        }

        [Test()]
        public void TestDescription()
        {
            var filePath = Path.Combine(_baseFilePath, "ProductWithChineseTitle.xlsx");
            var datas = _excelReader.ReadExcel(filePath, true);
            CheckDatas(datas);
        }

        [Test()]
        public void TestPropertyName()
        {
            var filePath = Path.Combine(_baseFilePath, "Product.xlsx");
            var datas = _excelReader.ReadExcel(filePath, false);
            CheckDatas(datas);
        }

        private void CheckDatas(List<Product> products)
        {
            Assert.AreEqual(products.Count, 3);

            var first = products[0];
            Assert.IsTrue(first.Name == "鉛筆" &&
                          first.Inventory == 10 &&
                          first.Unit == "枝" &&
                          first.Vender == "北方" &&
                          first.Price == 10 &&
                          first.Profit == 3);

            var second = products[1];
            Assert.IsTrue(second.Name == "尺" &&
                          second.Inventory == 5 &&
                          second.Unit == "把" &&
                          second.Vender == "夏風" &&
                          second.Price == 5 &&
                          second.Profit == 2);

            var thrid = products[2];
            Assert.IsTrue(thrid.Name == "足球" &&
                          thrid.Inventory == 3 &&
                          thrid.Unit == "顆" &&
                          thrid.Vender == "凱薩" &&
                          thrid.Price == 100 &&
                          thrid.Profit == 50);
        }
    }
}
