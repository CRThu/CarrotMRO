using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace CarrotMRO
{
    public class AppConfig
    {
        [YamlMember(Alias = "app")]
        public string AppName { get; set; }

        [YamlMember(Alias = "version")]
        public int Version { get; set; }

        [YamlMember(Alias = "standard")]
        public StandardConfig Standard { get; set; }

        [YamlMember(Alias = "quotation")]
        public QuotationConfig Quotation { get; set; }

        [YamlMember(Alias = "autosave")]
        public AutosaveConfig Autosave { get; set; }

        [YamlMember(Alias = "match")]
        public MatchConfig Match { get; set; }

        [YamlMember(Alias = "validate")]
        public ValidateConfig Validate { get; set; }

    }

    public class StandardConfig
    {
        [YamlMember(Alias = "name")]
        public string FileName { get; set; }

        [YamlMember(Alias = "header")]
        public HeaderConfig Header { get; set; }
    }

    public class QuotationConfig
    {
        [YamlMember(Alias = "group")]
        public bool Group { get; set; }

        [YamlMember(Alias = "header")]
        public HeaderConfig Header { get; set; }
    }

    public class AutosaveConfig
    {
        [YamlMember(Alias = "name")]
        public string FileName { get; set; }

        [YamlMember(Alias = "header")]
        public HeaderConfig Header { get; set; }
    }

    public class HeaderConfig
    {
        [YamlMember(Alias = "part")]
        public int? PartColumn { get; set; }

        [YamlMember(Alias = "customName")]
        public int? CustomNameColumn { get; set; }

        [YamlMember(Alias = "name")]
        public int? NameColumn { get; set; }

        [YamlMember(Alias = "unit")]
        public int? UnitColumn { get; set; }

        [YamlMember(Alias = "num")]
        public int? NumColumn { get; set; }

        [YamlMember(Alias = "perPrice")]
        public int? PerPriceColumn { get; set; }

        [YamlMember(Alias = "materialInPerPrice")]
        public int? MaterialInPerPriceColumn { get; set; }

        [YamlMember(Alias = "baseMaterialInPerPrice")]
        public int? BaseMaterialInPerPriceColumn { get; set; }

        [YamlMember(Alias = "auxMaterialInPerPrice")]
        public int? AuxMaterialInPerPriceColumn { get; set; }

        [YamlMember(Alias = "machinelInPerPrice")]
        public int? MachineInPerPriceColumn { get; set; }

        [YamlMember(Alias = "laborInPerPrice")]
        public int? LaborInPerPriceColumn { get; set; }

        //laborInPerPrice: 
        [YamlMember(Alias = "sumPrice")]
        public int? SumPriceColumn { get; set; }

        [YamlMember(Alias = "desc")]
        public int? DescColumn { get; set; }

        public List<DataGridColumnDefinition> BuildColumns()
        {
            var columns = new List<DataGridColumnDefinition>();

            // 直接逐个字段检查并映射
            if (PartColumn.HasValue)
                columns.Add(new() { Header = "类目", BindingPath = "Part" });

            if (CustomNameColumn.HasValue)
                columns.Add(new() { Header = "自定项目", BindingPath = "CustomName" });

            if (NameColumn.HasValue)
                columns.Add(new() { Header = "标准项目", BindingPath = "Name" });

            if (UnitColumn.HasValue)
                columns.Add(new() { Header = "单位", BindingPath = "Unit" });

            if (NumColumn.HasValue)
                columns.Add(new() { Header = "数量", BindingPath = "Num" });

            if (PerPriceColumn.HasValue)
                columns.Add(new() { Header = "综合单价", BindingPath = "PerPrice" });

            if (MaterialInPerPriceColumn.HasValue)
                columns.Add(new() { Header = "材料单价", BindingPath = "MaterialInPerPrice" });

            if (BaseMaterialInPerPriceColumn.HasValue)
                columns.Add(new() { Header = "主材单价", BindingPath = "BaseMaterialInPerPrice" });

            if (AuxMaterialInPerPriceColumn.HasValue)
                columns.Add(new() { Header = "辅材单价", BindingPath = "AuxMaterialInPerPrice" });

            if (MachineInPerPriceColumn.HasValue)
                columns.Add(new() { Header = "机械单价", BindingPath = "MachineInPerPrice" });

            if (LaborInPerPriceColumn.HasValue)
                columns.Add(new() { Header = "人工单价", BindingPath = "LaborInPerPrice" });

            if (SumPriceColumn.HasValue)
                columns.Add(new() { Header = "总价", BindingPath = "SumPrice", BindingMode = BindingMode.OneWay });

            if (DescColumn.HasValue)
                columns.Add(new() { Header = "备注", BindingPath = "Description" });

            return columns;
        }

        // 列名映射字典
        private static readonly Dictionary<string, string> _headerTexts = new()
        {
            [nameof(PartColumn)] = "类目",
            [nameof(CustomNameColumn)] = "自定项目",
            [nameof(NameColumn)] = "标准项目",
            [nameof(UnitColumn)] = "单位",
            [nameof(NumColumn)] = "数量",
            [nameof(PerPriceColumn)] = "综合单价",
            [nameof(MaterialInPerPriceColumn)] = "材料单价",
            [nameof(BaseMaterialInPerPriceColumn)] = "主材单价",
            [nameof(AuxMaterialInPerPriceColumn)] = "辅材单价",
            [nameof(MachineInPerPriceColumn)] = "机械单价",
            [nameof(LaborInPerPriceColumn)] = "人工单价",
            [nameof(SumPriceColumn)] = "总价",
            [nameof(DescColumn)] = "备注"
        };

        // 获取所有已配置的列（带列号）
        public Dictionary<int, string> GetHeaderMapping()
        {
            var mapping = new Dictionary<int, string>();
            var properties = GetType().GetProperties()
                .Where(p => p.Name.EndsWith("Column") && p.PropertyType == typeof(int?));

            foreach (var prop in properties)
            {
                var value = (int?)prop.GetValue(this);
                if (value.HasValue && _headerTexts.TryGetValue(prop.Name, out var headerText))
                {
                    mapping[value.Value] = headerText;
                }
            }
            return mapping;
        }

        // 生成完整表头数组（保留空位）
        public string[] GetFullHeaders()
        {
            var mapping = GetHeaderMapping();
            if (mapping.Count == 0) return Array.Empty<string>();

            int maxColumn = mapping.Keys.Max();
            var headers = new string[maxColumn];

            foreach (var kvp in mapping)
            {
                if (kvp.Key > 0)
                    headers[kvp.Key - 1] = kvp.Value; // 列号转为0-based索引
            }

            return headers;
        }
    }

    public class MatchConfig
    {
        [YamlMember(Alias = "unit")]
        public bool Unit { get; set; }

        [YamlMember(Alias = "num")]
        public bool Num { get; set; }

        [YamlMember(Alias = "perPrice")]
        public bool PerPrice { get; set; }

        [YamlMember(Alias = "desc")]
        public bool Desc { get; set; }
    }

    public class ValidateConfig
    {
        [YamlMember(Alias = "unit")]
        public bool Unit { get; set; }

        [YamlMember(Alias = "perPrice")]
        public bool PerPrice { get; set; }
    }


    public static class ConfigLoader
    {
        public static AppConfig LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Config file not found", filePath);

            var yamlContent = File.ReadAllText(filePath);
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            return deserializer.Deserialize<AppConfig>(yamlContent);
        }

        public static void SaveToFile(AppConfig config, string filePath)
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var yamlContent = serializer.Serialize(config);
            File.WriteAllText(filePath, yamlContent);
        }
    }
}