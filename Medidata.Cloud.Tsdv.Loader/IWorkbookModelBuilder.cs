namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IWorkbookModelBuilder
    {
        void AddWorksheet<TInterface>(string sheetName) where TInterface : class;
        IWorksheetModelBuilder GetWorksheet<TInterface>(string sheetName) where TInterface : class;
        void RemoveWorksheet(string sheetName);
        bool ContainsWorksheet(string sheetName);
        object ToModel();
    }
}