using System.IO;

namespace excelConvert.Services
{
    public interface IExportStrategy
    {
        void Export(object data, string filePath);
        string GetFileExtension();
    }
    
    public class JsonExportStrategy : IExportStrategy
    {
        public void Export(object data, string filePath)
        {
            string json = System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(filePath, json);
        }
        
        public string GetFileExtension()
        {
            return ".json";
        }
    }
    
    public class PbExportStrategy : IExportStrategy
    {
        public void Export(object data, string filePath)
        {
            string json = System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(filePath, json);
        }
        
        public string GetFileExtension()
        {
            return ".pb";
        }
    }
    
    public class ExportStrategyFactory
    {
        public static IExportStrategy CreateStrategy(string format)
        {
            switch (format.ToLower())
            {
                case "json":
                    return new JsonExportStrategy();
                case "pb":
                    return new PbExportStrategy();
                default:
                    return new JsonExportStrategy();
            }
        }
    }
}