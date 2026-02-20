using excelConvert.Services;
using System.IO;
using System;
using Xunit;

namespace excelConvert.Tests.Services
{
    public class ExportStrategyTests
    {
        [Fact]
        public void JsonExportStrategy_ShouldExportData_ToJsonFile()
        {
            // Arrange
            var exportStrategy = new JsonExportStrategy();
            var testData = new { Name = "Test", Value = 123 };
            string filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.json");
            
            // Act
            exportStrategy.Export(testData, filePath);
            
            // Assert
            Assert.True(File.Exists(filePath));
            var fileContent = File.ReadAllText(filePath);
            Assert.Contains("Test", fileContent);
            Assert.Contains("123", fileContent);
            
            // Clean up
            File.Delete(filePath);
        }
        
        [Fact]
        public void PbExportStrategy_ShouldExportData_ToPbFile()
        {
            // Arrange
            var exportStrategy = new PbExportStrategy();
            var testData = new { Name = "Test", Value = 123 };
            string filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.pb");
            
            // Act
            exportStrategy.Export(testData, filePath);
            
            // Assert
            Assert.True(File.Exists(filePath));
            var fileContent = File.ReadAllText(filePath);
            Assert.Contains("Test", fileContent);
            Assert.Contains("123", fileContent);
            
            // Clean up
            File.Delete(filePath);
        }
        
        [Fact]
        public void ExportStrategyFactory_ShouldCreateJsonStrategy_WhenFormatIsJson()
        {
            // Arrange & Act
            var strategy = ExportStrategyFactory.CreateStrategy("json");
            
            // Assert
            Assert.IsType<JsonExportStrategy>(strategy);
            Assert.Equal(".json", strategy.GetFileExtension());
        }
        
        [Fact]
        public void ExportStrategyFactory_ShouldCreatePbStrategy_WhenFormatIsPb()
        {
            // Arrange & Act
            var strategy = ExportStrategyFactory.CreateStrategy("pb");
            
            // Assert
            Assert.IsType<PbExportStrategy>(strategy);
            Assert.Equal(".pb", strategy.GetFileExtension());
        }
        
        [Fact]
        public void ExportStrategyFactory_ShouldCreateJsonStrategy_WhenFormatIsInvalid()
        {
            // Arrange & Act
            var strategy = ExportStrategyFactory.CreateStrategy("invalid");
            
            // Assert
            Assert.IsType<JsonExportStrategy>(strategy);
            Assert.Equal(".json", strategy.GetFileExtension());
        }
    }
}