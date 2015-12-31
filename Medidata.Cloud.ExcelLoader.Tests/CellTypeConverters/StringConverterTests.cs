using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ploeh.AutoFixture;
using Ploeh.AutoFixture.AutoRhinoMock;

namespace Medidata.Cloud.ExcelLoader.Tests.CellTypeConverters
{
    [TestClass]
    public class StringConverterTests
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
            var converter = new StringConverter();
            object value0;
            object value1;
            object value2;
            object value3;
            //Act
            bool success0 = converter.TryConvertToCSharpValue("0", out value0);
            bool success1 = converter.TryConvertToCSharpValue("1", out value1);
            bool success2 = converter.TryConvertToCSharpValue(null, out value2);
            bool success3 = converter.TryConvertToCSharpValue(string.Empty, out value3);

            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.IsTrue(success3);
            Assert.AreEqual("0", value0);
            Assert.AreEqual("1",value1);
            Assert.IsNull(value2);
            Assert.AreEqual(string.Empty,value3);

        }


        [TestMethod]
        public void ShouldConvertToCellValue()
        {
            //Arrange

            //Act
            var converter = new StringConverter();
            string value0;
            string value1;
            bool success0 = converter.TryConvertToCellValue("Yes", out value0);
            bool success1 = converter.TryConvertToCellValue("No", out value1);

            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.AreEqual("Yes", value0);
            Assert.AreEqual("No", value1);
        }


        [TestMethod]
        public void ShouldFailToConvertToCellValue()
        {
            //Arrange

            //Act
            var converter = new StringConverter();
            string value0;
            string value1;
            bool success0 = converter.TryConvertToCellValue(0, out value0);
            bool success1 = converter.TryConvertToCellValue(9E+09, out value1);

            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
        }
    }
}