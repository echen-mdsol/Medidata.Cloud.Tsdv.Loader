using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Medidata.Cloud.Tsdv.Loader.ViewModels;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Medidata.Cloud.Tsdv.Loader.Tests
{
    [TestClass]
    public class DemoTest
    {
        [TestMethod]
        public void Demo1()
        {
            var blockPlans = new List<BlockPlan>
            {
                new BlockPlan
                {
                    Activated = true,
                    ActivatedUserName = "cheng",
                    AverageSubjectPerSite = 15,
                    BlockPlanType = "Study",
                    CoveragePercent = 95,
                    DateEstimated = DateTime.Now,
                    IsProdInUse = true,
                    MatrixName = "DefaultMatrix",
                    Name = "Plan A",
                    ObjectName = "ObjectName1",
                    RoleName = "Role 1"
                }
            };
            var result = blockPlans.ToExcel(@"TestHelpers\TSDV.xslt");
            Debug.WriteLine(result);
            File.WriteAllText(@"c:\temp\blockplan.xml",result);
        }
    }
}