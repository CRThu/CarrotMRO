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
        [NotifyPropertyChangedFor(nameof(SumPrice))]
        private double num = 1;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(SumPrice))]
        private double perPrice = 0;

        [ObservableProperty]
        private string description = "";

        public double SumPrice => Num * PerPrice;
    }
}
