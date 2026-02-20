using excelConvert.Models;
using System.Collections.Generic;

namespace excelConvert.Services
{
    public interface IConfigFactory
    {
        Config CreateConfig();
        Config CreateConfig(string jsonContent);
    }
    
    public class ConfigFactory : IConfigFactory
    {
        public Config CreateConfig()
        {
            return new Config
            {
                Sheet1 = new SheetConfig
                {
                    DataTypes = new List<DataTypeConfig>
                    {
                        new DataTypeConfig
                        {
                            Name = "Sheet1",
                            Description = "Default sheet",
                            Fields = new List<FieldConfig>(),
                            ExportFormats = new List<string> { "json", "xml" }
                        }
                    }
                }
            };
        }
        
        public Config CreateConfig(string jsonContent)
        {
            try
            {
                var config = System.Text.Json.JsonSerializer.Deserialize<Config>(jsonContent);
                return config ?? CreateConfig();
            }
            catch
            {
                return CreateConfig();
            }
        }
    }
}