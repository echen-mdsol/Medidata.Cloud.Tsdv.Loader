using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Builders;
using Medidata.Cloud.Tsdv.Loader.Tests.Helpers;
using Medidata.Interfaces.TSDV;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using YAXLib;

namespace Medidata.Cloud.Tsdv.Loader.Tests
{
    [TestClass]
    public class SpreadsheetBuilderDemo
    {
        [TestMethod]
        public void CreateSpreadsheet()
        {
            var ssBuilder = new SpreadsheetBuilder();

            var sheet = ssBuilder.EnsureWorksheet<IFakeInterface>("MySheet1");
            sheet.Add(new FakeInterfaceClass { Name = "xxx", DateOfMarriage = null, IsAdult = false});
            sheet.Add(new FakeInterfaceClass { Name = "yyy" });
            sheet.Add(new FakeInterfaceClass { Name = "zzz" });

            var sheet2 = ssBuilder.EnsureWorksheet<IFakeInterface>("MySheet2", new[] { "tsdv_Col1", "tsdv_Col2" });
            sheet2.Add(new FakeInterfaceClass { Name = "qqq", Birthday = DateTime.MaxValue, Age = 34, Height = 1.5});
            sheet2.Add(new FakeInterfaceClass { Name = "111", Salary = (decimal) 1234.324});
            sheet2.Add(new FakeInterfaceClass { Name = "ccc", DateOfMarriage = DateTime.Now});

            var filePath = @"C:\Github\test.xlsx";
            File.Delete(filePath);
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                var doc = ssBuilder.Save(fs);
            }

            Assert.IsTrue(File.Exists(filePath));
        }
    }
}
