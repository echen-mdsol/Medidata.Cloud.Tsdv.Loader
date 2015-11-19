using System;

namespace Medidata.Cloud.Tsdv.Loader.Tests.Helpers
{
    public interface IFakeInterface
    {
        string Name { get; }
        decimal Salary { get; }
        int Age { get; }
        double Height { get; }
        DateTime Birthday { get; }
        DateTime? DateOfMarriage { get; }
        bool IsAdult { get; }
    }
}