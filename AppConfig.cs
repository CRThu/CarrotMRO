using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        [YamlMember(Alias = "autosave")]
        public AutosaveConfig Autosave { get; set; }
    }

    public class StandardConfig
    {
        [YamlMember(Alias = "name")]
        public string FileName { get; set; }

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
                columns.Add(new() { Header = "单价", BindingPath = "PerPrice" });

            if (DescColumn.HasValue)
                columns.Add(new() { Header = "备注", BindingPath = "Description" });

            return columns;
        }
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