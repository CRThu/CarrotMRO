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