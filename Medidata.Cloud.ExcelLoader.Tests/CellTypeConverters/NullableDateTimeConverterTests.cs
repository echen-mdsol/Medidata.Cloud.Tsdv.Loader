using System;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ploeh.AutoFixture;
using Ploeh.AutoFixture.AutoRhinoMock;

namespace Medidata.Cloud.ExcelLoader.Tests.CellTypeConverters
{
    [TestClass]
    public class NullableDateTimeConverterTests
    {
        private IFixture _fixture;
        [TestInitialize]
        public void Init()
        {
            _fixture = new Fixture().Customize(new AutoRhinoMockCustomization());
        }

        [TestMethod]
        public void ShouldConvertToCSharpValue()
        {
            //Arrange
            var converter = new NullableDateTimeConverter();
            object value0;
            object value1;
            object value2;
            object value3;
            //Act
            bool success0 = converter.TryConvertToCSharpValue("2015-12-15", out value0);
            bool success1 = converter.TryConvertToCSharpValue("12/15/2015", out value1);
            bool success2 = converter.TryConvertToCSharpValue("12-15-2015", out value2);
            bool success3 = converter.TryConvertToCSharpValue("", out value3);
            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.IsTrue(success3);
            var targetDate = new DateTime(2015, 12, 15);
            Assert.AreEqual(targetDate, value0);
            Assert.AreEqual(targetDate, value1);
            Assert.AreEqual(targetDate, value2);
            Assert.IsNull(value3);


        }


        [TestMethod]
        public void ShouldFailToConvertToCSharpValue()
        {
            //Arrange
            var converter = new NullableDateTimeConverter();
            object value0;
            object value1;
            object value2;
            //Act
            bool success0 = converter.TryConvertToCSharpValue("2/29/2015", out value0);
            bool success1 = converter.TryConvertToCSharpValue("15/12/2015", out value1);
            bool success2 = converter.TryConvertToCSharpValue("NoDate", out value2);
            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
            Assert.IsFalse(success2);


        }

        [TestMethod]
        public void ShouldConvertToCellValue()
        {
            //Arrange
            var converter = new NullableDateTimeConverter();
            string value0;
            string value1;
            string value2;
            string value3;
            //Act
            bool success0 = converter.TryConvertToCellValue(new DateTime(2015, 10, 10), out value0);
            bool success1 = converter.TryConvertToCellValue(DateTime.MinValue, out value1);
            bool success2 = converter.TryConvertToCellValue(DateTime.MaxValue, out value2);
            bool success3 = converter.TryConvertToCellValue(null, out value3);
            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.IsTrue(success3);
            Assert.AreEqual("2015/10/10", value0);
            Assert.AreEqual("0001/01/01", value1);
            Assert.AreEqual("9999/12/31", value2);
            Assert.AreEqual(null,value3);


        }

        [TestMethod]
        public void ShouldFailToConvertToCellValue()
        {
            //Arrange
            var converter = new NullableDateTimeConverter();
            string value0;
            string value1;
            string value2;
            //Act
            bool success0 = converter.TryConvertToCellValue("10/25/2015", out value0);
            bool success1 = converter.TryConvertToCellValue("N/A", out value1);
            bool success2 = converter.TryConvertToCellValue("", out value2);
            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
            Assert.IsFalse(success2);
        }
    }
}