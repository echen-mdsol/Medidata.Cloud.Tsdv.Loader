using System;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class IntConverter : NumberConverter<int>
    {
        protected override int GetCSharpValueImpl(string cellValue)
        {
            return int.Parse(cellValue);
        }
    }
}