using System;
using ImpromptuInterface;
using Medidata.Cloud.Tsdv.Loader.Models;
using Medidata.Interfaces.TSDV;

namespace Medidata.Cloud.Tsdv.Loader.ModelConverters
{
    public class BlockPlanConverter : IModelConverter
    {
        public BlockPlanConverter()
        {
            InterfaceType = typeof (IBlockPlan);
            ViewModelType = typeof (BlockPlanModel);
        }

        public object ConvertToViewModel(object target)
        {
            throw new NotImplementedException();
        }

        public object ConvertFromViewModel(object viewModel)
        {
            IBlockPlan target = viewModel.ActLike();
            return target;
        }

        public Type InterfaceType { get; private set; }
        public Type ViewModelType { get; private set; }
    }
}