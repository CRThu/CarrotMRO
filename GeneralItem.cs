using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarrotMRO
{
    public partial class GeneralItem : ObservableObject
    {
        [ObservableProperty]
        private string part = "<默认类目>";

        [ObservableProperty]
        private string customName = "<自定项目名>";

        [ObservableProperty]
        private string name = "<标准项目名>";

        [ObservableProperty]
        private string unit = "<计算单位>";

        [ObservableProperty]
        private double num = 1;

        [ObservableProperty]
        private double? perPrice;

        [ObservableProperty]
        private double? materialInPerPrice;

        [ObservableProperty]
        private double? baseMaterialInPerPrice;

        [ObservableProperty]
        private double? auxMaterialInPerPrice;

        [ObservableProperty]
        private double? machineInPerPrice;

        [ObservableProperty]
        private double? laborInPerPrice;

        [ObservableProperty]
        private string description = "";

        [ObservableProperty]
        public double sumPrice;


        // 当“数量”变化时，直接更新总价
        partial void OnNumChanged(double value)
        {
            UpdateSumPrice();
        }

        // 当“主材单价”变化时，触发“材料单价”的更新
        partial void OnBaseMaterialInPerPriceChanged(double? value)
        {
            UpdateMaterialInPerPrice();
        }

        // 当“辅材单价”变化时，同样触发“材料单价”的更新
        partial void OnAuxMaterialInPerPriceChanged(double? value)
        {
            UpdateMaterialInPerPrice();
        }

        // 当“材料单价”变化时（无论是手动修改还是计算得出），触发“综合单价”的更新
        partial void OnMaterialInPerPriceChanged(double? value)
        {
            UpdatePerPrice();
        }

        // 当“机械单价”变化时，触发“综合单价”的更新
        partial void OnMachineInPerPriceChanged(double? value)
        {
            UpdatePerPrice();
        }

        // 当“人工单价”变化时，触发“综合单价”的更新
        partial void OnLaborInPerPriceChanged(double? value)
        {
            UpdatePerPrice();
        }

        // 当“综合单价”变化时（无论是手动修改还是计算得出），触发“总价”的更新
        partial void OnPerPriceChanged(double? value)
        {
            UpdateSumPrice();
        }

        /// <summary>
        /// 根据主材和辅材单价，计算并更新材料单价
        /// </summary>
        private void UpdateMaterialInPerPrice()
        {
            // 使用 ?? 0 来处理可空类型，如果为 null 则按 0 计算
            MaterialInPerPrice = (BaseMaterialInPerPrice ?? 0) + (AuxMaterialInPerPrice ?? 0);
        }

        /// <summary>
        /// 根据材料、机械、人工单价，计算并更新综合单价
        /// </summary>
        private void UpdatePerPrice()
        {
            PerPrice = (MaterialInPerPrice ?? 0) + (MachineInPerPrice ?? 0) + (LaborInPerPrice ?? 0);
        }

        /// <summary>
        /// 根据数量和综合单价，计算并更新总价
        /// </summary>
        private void UpdateSumPrice()
        {
            SumPrice = Num * (PerPrice ?? 0);
        }
    }
}
