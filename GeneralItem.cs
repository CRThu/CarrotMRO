using CommunityToolkit.Mvvm.ComponentModel;

namespace CarrotMRO
{
    public partial class GeneralItem : ObservableObject
    {
        // 一个内部标志，用于防止在计算更新时触发无限循环。
        // 当我们通过代码（而不是用户输入）来设置一个属性时，我们会将此标志设为true。
        private bool _isUpdatingInternally = false;

        [ObservableProperty]
        private string part = "<默认类目>";

        [ObservableProperty]
        private string customName = "<自定项目名>";

        [ObservableProperty]
        private string name = "<标准项目名>";

        [ObservableProperty]
        private string unit = "<计算单位>";

        [ObservableProperty]
        private decimal num = 1;

        [ObservableProperty]
        private decimal? perPrice;

        [ObservableProperty]
        private decimal? materialInPerPrice;

        [ObservableProperty]
        private decimal? baseMaterialInPerPrice;

        [ObservableProperty]
        private decimal? auxMaterialInPerPrice;

        [ObservableProperty]
        private decimal? machineInPerPrice;

        [ObservableProperty]
        private decimal? laborInPerPrice;

        [ObservableProperty]
        private string description = "";

        [ObservableProperty]
        public decimal sumPrice;

        public void SetInitialValues(
            string part,
            string customName,
            string name,
            string unit,
            decimal num,
            decimal? perPrice,
            decimal? materialInPerPrice,
            decimal? baseMaterialInPerPrice,
            decimal? auxMaterialInPerPrice,
            decimal? machineInPerPrice,
            decimal? laborInPerPrice,
            string description,
            decimal sumPrice
            )
        {
            // 步骤1：开启“内部更新”模式，禁用所有联动逻辑
            _isUpdatingInternally = true;

            // 步骤2：直接、安全地为所有属性赋值
            if (part != null) this.Part = part;
            if (customName != null) this.CustomName = customName;
            if (name != null) this.Name = name;
            if (unit != null) this.Unit = unit;
            this.Num = num;
            this.PerPrice = perPrice;
            this.MaterialInPerPrice = materialInPerPrice;
            this.BaseMaterialInPerPrice = baseMaterialInPerPrice;
            this.AuxMaterialInPerPrice = auxMaterialInPerPrice;
            this.MachineInPerPrice = machineInPerPrice;
            this.LaborInPerPrice = laborInPerPrice;
            if(description != null) this.Description = description;
            this.SumPrice = sumPrice;

            // 步骤3：关闭“内部更新”模式，恢复正常的交互逻辑
            _isUpdatingInternally = false;

            // 步骤4：最后，根据已加载的数据，强制从上到下重新计算一次总价，以确保显示正确
            UpdateMaterialInPerPrice();
        }

        public GeneralItem Copy()
        {
            // MemberwiseClone() 会创建一个浅拷贝。
            // 因为我们所有的字段都是值类型(decimal)或不可变类型(string)，
            // 所以浅拷贝的效果等同于深拷贝，是完全安全的。
            // 返回的对象是一个全新的实例，与原对象完全独立。
            return (GeneralItem)this.MemberwiseClone();
        }

        // 数量变化，直接影响最终总价
        partial void OnNumChanged(decimal value)
        {
            UpdateSumPrice();
        }

        // 最底层的分项变化时，总是触发向上的更新
        partial void OnBaseMaterialInPerPriceChanged(decimal? value)
        {
            // 如果是内部计算导致的变更，则不执行任何操作，避免循环
            if (_isUpdatingInternally) return;
            UpdateMaterialInPerPrice();
        }

        partial void OnAuxMaterialInPerPriceChanged(decimal? value)
        {
            if (_isUpdatingInternally) return;
            UpdateMaterialInPerPrice();
        }

        partial void OnMachineInPerPriceChanged(decimal? value)
        {
            if (_isUpdatingInternally) return;
            UpdatePerPrice();
        }

        partial void OnLaborInPerPriceChanged(decimal? value)
        {
            if (_isUpdatingInternally) return;
            UpdatePerPrice();
        }

        // 当材料单价被手动修改时
        partial void OnMaterialInPerPriceChanged(decimal? value)
        {
            if (_isUpdatingInternally) return;

            // 用户手动修改了材料单价，检查是否需要清空下级
            if (value.HasValue)
            {
                var calculatedFromChildren = (BaseMaterialInPerPrice ?? 0) + (AuxMaterialInPerPrice ?? 0);
                // 如果输入的值和下级计算的值不相等，则清空下级
                if (value.Value != calculatedFromChildren)
                {
                    _isUpdatingInternally = true;
                    BaseMaterialInPerPrice = null;
                    AuxMaterialInPerPrice = null;
                    _isUpdatingInternally = false;
                }
            }
            // 无论如何，材料单价的变化都会影响综合单价
            UpdatePerPrice();
        }

        // 当综合单价被手动修改时
        partial void OnPerPriceChanged(decimal? value)
        {
            if (_isUpdatingInternally) return;

            // 用户手动修改了综合单价，检查是否需要清空下级
            if (value.HasValue)
            {
                var calculatedFromChildren = (MaterialInPerPrice ?? 0) + (MachineInPerPrice ?? 0) + (LaborInPerPrice ?? 0);
                // 如果输入的值和下级计算的值不相等，则清空下级
                if (value.Value != calculatedFromChildren)
                {
                    _isUpdatingInternally = true;
                    MaterialInPerPrice = null;
                    MachineInPerPrice = null;
                    LaborInPerPrice = null;
                    // 由于MaterialInPerPrice被清空，它的下级也应被清空
                    BaseMaterialInPerPrice = null;
                    AuxMaterialInPerPrice = null;
                    _isUpdatingInternally = false;
                }
            }

            // 无论如何，综合单价的变化都会影响总价
            UpdateSumPrice();
        }


        /// <summary>
        /// 根据主材和辅材单价，更新材料单价
        /// </summary>
        private void UpdateMaterialInPerPrice()
        {
            _isUpdatingInternally = true;
            // 如果下级都为null，则自身也为null；否则进行计算
            if (!BaseMaterialInPerPrice.HasValue && !AuxMaterialInPerPrice.HasValue)
            {
                MaterialInPerPrice = null;
            }
            else
            {
                MaterialInPerPrice = (BaseMaterialInPerPrice ?? 0) + (AuxMaterialInPerPrice ?? 0);
            }
            _isUpdatingInternally = false;

            // 链式更新：材料单价更新后，需要继续更新综合单价
            UpdatePerPrice();
        }

        /// <summary>
        /// 根据材料、机械、人工单价，更新综合单价
        /// </summary>
        private void UpdatePerPrice()
        {
            _isUpdatingInternally = true;
            // 如果下级都为null，则自身也为null；否则进行计算
            if (!MaterialInPerPrice.HasValue && !MachineInPerPrice.HasValue && !LaborInPerPrice.HasValue)
            {
                PerPrice = null;
            }
            else
            {
                PerPrice = (MaterialInPerPrice ?? 0) + (MachineInPerPrice ?? 0) + (LaborInPerPrice ?? 0);
            }
            _isUpdatingInternally = false;

            // 链式更新：综合单价更新后，需要继续更新总价
            UpdateSumPrice();
        }

        /// <summary>
        /// 根据数量和综合单价，更新总价
        /// </summary>
        private void UpdateSumPrice()
        {
            // SumPrice 是最终值，不是可空的，所以这里 ?? 0 是合适的
            SumPrice = Num * (PerPrice ?? 0);
        }
    }
}