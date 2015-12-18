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
        private readonly IDictionary<Type, SheetInfo> _sheetInfoDic = new Dictionary<Type, SheetInfo>();
        private readonly ISheetParser _sheetParser;

        public ExcelLoader(IExcelBuilder builder, IExcelParser parser, ISheetBuilder sheetBuilder,
                           ISheetParser sheetParser)
        {
            _sheetBuilder = sheetBuilder;
            _sheetParser = sheetParser;

            _builder = builder;
            _parser = parser;
        }

        public virtual void Save(Stream outStream)
        {
            foreach (var type in _sheetInfoDic.Keys)
            {
                var info = _sheetInfoDic[type];
                var sheetDef = info.SheetDefinition;
                var sheetData = info.DataForSave ?? new List<SheetModel>();
                _builder.AddSheet(sheetDef, sheetData.Cast<SheetModel>(), _sheetBuilder);
            }

            _builder.Save(outStream);
        }

        public virtual void Load(Stream source)
        {
            _parser.Load(source);

            foreach (var type in _sheetInfoDic.Keys)
            {
                var info = _sheetInfoDic[type];
                var sheetDef = info.SheetDefinition;
                info.LoadedData = _parser.GetObjects(sheetDef, _sheetParser).ToList();
            }
        }

        public ISheetInfo<T> Sheet<T>() where T : SheetModel
        {
            SheetInfo info;
            if (!_sheetInfoDic.TryGetValue(typeof(T), out info))
            {
                var sheetDef = SheetDefinition.Define<T>();
                info = new SheetInfo {SheetDefinition = sheetDef};
                _sheetInfoDic.Add(typeof(T), info);
            }
            return new SheetInfo<T>(info);
        }

        private class SheetInfo
        {
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