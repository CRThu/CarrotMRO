using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace CarrotMRO
{
    public class DataGridColumnDefinition
    {
        public string Header { get; set; }
        public string BindingPath { get; set; }
        public BindingMode BindingMode { get; set; } = BindingMode.TwoWay;
    }
}
