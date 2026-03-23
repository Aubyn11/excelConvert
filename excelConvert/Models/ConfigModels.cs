using System.Collections.Generic;

namespace excelConvert.Models
{
    public class Config
    {
        public SheetConfig Sheet1 { get; set; }
    }

    public class SheetConfig
    {
        public List<DataTypeConfig> DataTypes { get; set; }
    }

    public class DataTypeConfig
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public List<FieldConfig> Fields { get; set; }
        public List<string> ExportFormats { get; set; }
    }

    public class FieldConfig
    {
        public string ExportName { get; set; }
        public string ExcelColumn { get; set; }
        public bool Required { get; set; }
        public string Type { get; set; }
        public bool IsRepeated { get; set; }
    }
}
