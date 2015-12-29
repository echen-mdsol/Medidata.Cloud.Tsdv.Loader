using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ploeh.AutoFixture;
using Ploeh.AutoFixture.AutoRhinoMock;

namespace Medidata.Cloud.ExcelLoader.Tests.CellTypeConverters
{
    [TestClass]
    public class DoubleConverterTests
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
            object value0;
            object value1;
            object value2;
            var converter = new DoubleConverter();
            //Act
            bool success0 = converter.TryConvertToCSharpValue("0.123", out value0);
            bool success1 = converter.TryConvertToCSharpValue("9E10", out value1);
            bool success2 = converter.TryConvertToCSharpValue("-12.84848715", out value2);
            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.AreEqual(0.123, value0);
            Assert.AreEqual(9E10, value1);
            Assert.AreEqual(-12.84848715, value2);
        

        }



        [TestMethod]
        public void ShouldFailToConvertToCSharpValue()
        {
            //Arrange
            object value0;
            object value1;
            object value2;
            var converter = new DoubleConverter();
            //Act
            bool success0 = converter.TryConvertToCSharpValue("9E10D", out value0);
            bool success1 = converter.TryConvertToCSharpValue("N/A", out value1);
            bool success2 = converter.TryConvertToCSharpValue("10000E", out value2);
            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
            Assert.IsFalse(success2);
        }

        [TestMethod]
        public void ShouldConvertToCellValue()
        {
            //Arrange
            string value0;
            string value1;
            string value2;
            var converter = new DoubleConverter();
            //Act
            bool success0 = converter.TryConvertToCellValue(0.123D, out value0);
            bool success1 = converter.TryConvertToCellValue(1000000000D, out value1);
            bool success2 = converter.TryConvertToCellValue(-12.84848715D, out value2);
            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.AreEqual("0.123", value0);
            Assert.AreEqual("1000000000", value1);
            Assert.AreEqual("-12.84848715", value2);

        }

        [TestMethod]
        public void ShouldFailToConvertToCellValue()
        {
            //Arrange
            string value0;
            string value1;
            string value2;
            var converter = new DoubleConverter();
            //Act
            bool success0 = converter.TryConvertToCellValue(0.18f, out value0);
            bool success1 = converter.TryConvertToCellValue("-1", out value1);
            bool success2 = converter.TryConvertToCellValue("NaN", out value2);
            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
            Assert.IsFalse(success2);
        }


    }
}