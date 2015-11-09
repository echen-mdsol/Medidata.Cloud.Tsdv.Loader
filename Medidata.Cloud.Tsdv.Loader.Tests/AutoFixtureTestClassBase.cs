using Ploeh.AutoFixture;
using Ploeh.AutoFixture.AutoRhinoMock;
using Rhino.Mocks;

namespace Medidata.Cloud.Tsdv.Loader.Tests
{
    public abstract class AutoFixtureTestClassBase<T> where T : class
    {
        protected AutoFixtureTestClassBase()
        {
            Fixture = new Fixture().Customize(new AutoRhinoMockCustomization());
            Target = MockRepository.GeneratePartialMock<T>();
        }

        protected IFixture Fixture { get; private set; }
        protected T Target { get; set; }
    }
}