using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using Medidata.Cloud.ExcelLoader.Helpers;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelLoader : IExcelLoader
    {
        private readonly IExcelBuilder _builder;
        private readonly IExcelParser _parser;
        private readonly ISheetBuilder _sheetBuilder;
        private readonly IList<SheetInfo> _sheetInfos = new List<SheetInfo>();
        private readonly ISheetParser _sheetParser;

        public ExcelLoader(IExcelBuilder builder, IExcelParser parser, ISheetBuilder sheetBuilder,
                           ISheetParser sheetParser)
        {
            if (builder == null) throw new ArgumentNullException("builder");
            if (parser == null) throw new ArgumentNullException("parser");
            if (sheetBuilder == null) throw new ArgumentNullException("sheetBuilder");
            if (sheetParser == null) throw new ArgumentNullException("sheetParser");
            _sheetBuilder = sheetBuilder;
            _sheetParser = sheetParser;
            _builder = builder;
            _parser = parser;
        }

        public virtual void Save(Stream outStream)
        {
            if (outStream == null) throw new ArgumentNullException("outStream");
            foreach (var info in _sheetInfos)
            {
                var sheetDef = info.SheetDefinition;
                var sheetData = info.DataForSave ?? new List<SheetModel>();
                _builder.AddSheet(sheetDef, sheetData.Cast<SheetModel>(), _sheetBuilder);
            }

            _builder.Save(outStream);
        }

        public virtual void Load(Stream source)
        {
            if (source == null) throw new ArgumentNullException("source");
            _parser.Load(source);

            foreach (var info in _sheetInfos)
            {
                var sheetDef = info.SheetDefinition;
                info.LoadedData = _parser.GetObjects(sheetDef, _sheetParser).ToList();
            }
        }

        public virtual ISheetInfo<T> Sheet<T>() where T : SheetModel
        {
            var type = typeof(T);
            var info = _sheetInfos.FirstOrDefault(x => x.Type == type);
            if (info != null) return new SheetInfo<T>(info);
            var sheetDef = SheetDefinition.Define<T>();
            info = new SheetInfo {SheetDefinition = sheetDef, Type = type};
            _sheetInfos.Add(info);
            return new SheetInfo<T>(info);
        }

        private class SheetInfo
        {
            public Type Type { get; set; }
            public ISheetDefinition SheetDefinition { get; set; }
            public IList DataForSave { get; set; }
            public IList<ExpandoObject> LoadedData { get; set; }
        }

        private class SheetInfo<T> : ISheetInfo<T> where T : SheetModel
        {
            private readonly SheetInfo _info;

            public SheetInfo(SheetInfo info)
            {
                _info = info;
            }

            public IList<T> Data
            {
                get
                {
                    if (_info.LoadedData != null)
                    {
                        _info.DataForSave = _info.LoadedData.CastToSheetModel<T>().ToList();
                    }
                    else if (_info.DataForSave == null)
                    {
                        _info.DataForSave = new List<T>();
                    }
                    return (IList<T>) _info.DataForSave;
                }
            }

            public ISheetDefinition Definition
            {
                get { return _info.SheetDefinition; }
            }
        }
    }
}