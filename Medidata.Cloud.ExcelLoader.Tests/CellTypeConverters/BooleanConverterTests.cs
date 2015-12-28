using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ploeh.AutoFixture;
using Ploeh.AutoFixture.AutoRhinoMock;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
namespace Medidata.Cloud.ExcelLoader.Tests.CellTypeConverters
{
    [TestClass]
    public class BooleanConverterTests
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
            var converter = new BooleanConverter();
            object value0;
            object value1;
            object value2;
            object value3;
            //Act
            bool success0 = converter.TryConvertToCSharpValue("0", out value0);
            bool success1 = converter.TryConvertToCSharpValue("1", out value1);
            bool success2 = converter.TryConvertToCSharpValue("False", out value2);
            bool success3 = converter.TryConvertToCSharpValue("True", out value3);

            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.IsTrue(success3);
            Assert.IsFalse((bool)value0);
            Assert.IsTrue((bool)value1);
            Assert.IsFalse((bool)value2);
            Assert.IsTrue((bool)value3);

        }


        [TestMethod]
        public void ShouldFailToConvertToCSharpValue()
        {
            //Arrange

            //Act
            var converter = new BooleanConverter();
            object value0;
            object value1;
            bool success0 = converter.TryConvertToCSharpValue("666", out value0);
            bool success1 = converter.TryConvertToCSharpValue("Empty", out value1);

            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
        }

        [TestMethod]
        public void ShouldConvertToCellValue()
        {
            //Arrange

            //Act
            var converter = new BooleanConverter();
            string value0;
            string value1;
            bool success0 = converter.TryConvertToCellValue(false, out value0);
            bool success1 = converter.TryConvertToCellValue(true, out value1);

            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.AreEqual("0", value0);
            Assert.AreEqual("1", value1);
        }


        [TestMethod]
        public void ShouldFailToConvertToCellValue()
        {
            //Arrange

            //Act
            var converter = new BooleanConverter();
            string value0;
            string value1;
            bool success0 = converter.TryConvertToCellValue("True", out value0);
            bool success1 = converter.TryConvertToCellValue("bbb", out value1);

            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
        }
    }
}