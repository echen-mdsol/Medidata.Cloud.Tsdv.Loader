using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Medidata.Cloud.Tsdv.Loader.Builders;
using Medidata.Cloud.Tsdv.Loader.Parsers;
using Medidata.Cloud.Tsdv.Loader.Tests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Medidata.Cloud.Tsdv.Loader.Tests
{
    [TestClass]
    public class ExcelBuilderParserDemo
    {
        [TestMethod]
        public void CreateSpreadsheet()
        {
            var builder = new ExcelBuilder();

            var sheet = builder.AddSheet<IFakeInterface>("MySheet1", true);
            sheet.Add(new {Name = "xxx", IsAdult = false});
            sheet.Add(new {Name = "yyy"});
            sheet.Add(new {Name = "zzz"});

            var sheet2 = builder.AddSheet<IFakeInterface>("MySheet2", new[] {"tsdv_Col1", "tsdv_Col2"});
            sheet2.Add(new {Name = "qqq", Birthday = DateTime.MaxValue, Age = 34, Height = 1.5});
            sheet2.Add(new FakeInterfaceClass {Name = "111", Salary = (decimal) 1234.324});
            sheet2.Add(new FakeInterfaceClass {Name = "ccc", DateOfMarriage = DateTime.Now});

            var filePath = @"C:\Github\test.xlsx";
            File.Delete(filePath);
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                builder.Save(fs);
            }

            Assert.IsTrue(File.Exists(filePath));
        }

        [TestMethod]
        public void LoadSpreadsheet()
        {
            var filePath = @"C:\Github\test.xlsx";

            IList<IFakeInterface> objectsFromSheet1;
            IList<IFakeInterface> objectsFromSheet2;
            using (var fs = new FileStream(filePath, FileMode.Open))
            using (var parser = new ExcelParser())
            {
                parser.Load(fs);

                objectsFromSheet1 = parser.GetObjects<IFakeInterface>("MySheet1").ToList();

                objectsFromSheet2 = parser.GetObjects<IFakeInterface>("MySheet2").ToList();
            }

            Assert.IsNotNull(objectsFromSheet1);
            Assert.IsNotNull(objectsFromSheet2);
        }
    }
}