﻿using System;
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
    public class DecimalConverterTests
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
            var converter = new DecimalConverter();
            //Act
            bool success0 = converter.TryConvertToCSharpValue("0.123", out value0);
            bool success1 = converter.TryConvertToCSharpValue("10000000000", out value1);
            bool success2 = converter.TryConvertToCSharpValue("-12.84848715", out value2);
            //Assert
            Assert.IsTrue(success0);
            Assert.IsTrue(success1);
            Assert.IsTrue(success2);
            Assert.AreEqual(0.123m, value0);
            Assert.AreEqual(10000000000m, value1);
            Assert.AreEqual(-12.84848715m, value2);


        }



        [TestMethod]
        public void ShouldFailToConvertToCSharpValue()
        {
            //Arrange
            object value0;
            object value1;
            object value2;
            var converter = new DecimalConverter();
            //Act
            bool success0 = converter.TryConvertToCSharpValue("9E10", out value0);
            bool success1 = converter.TryConvertToCSharpValue("1000000000000000000000000000000000000000000000000000000000", out value1);
            bool success2 = converter.TryConvertToCSharpValue("0/0", out value2);
            //Assert
            Assert.IsFalse(success0);
            Assert.IsFalse(success1);
            Assert.IsFalse(success2);
        }
    }
}
