using System;

namespace Medidata.Cloud.Tsdv.Loader.Tests.Helpers
{
    public class FakeInterfaceClass : IFakeInterface
    {
        public string Name { get; set; }
        public decimal Salary { get; set; }
        public int Age { get; set; }
        public double Height { get; set; }
        public DateTime Birthday { get; set; }
        public DateTime? DateOfMarriage { get; set; }
        public bool IsAdult { get; set; }
    }
}