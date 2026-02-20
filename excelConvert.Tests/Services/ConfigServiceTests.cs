using excelConvert.Services;
using Xunit;

namespace excelConvert.Tests.Services
{
    public class ConfigServiceTests
    {
        [Fact]
        public void LoadConfig_ShouldReturnFalse_WhenFileDoesNotExist()
        {
            // Arrange
            var configService = new ConfigService();
            string nonExistentFile = "non_existent_file.json";
            
            // Act
            var result = configService.LoadConfig(nonExistentFile, out var config);
            
            // Assert
            Assert.False(result);
            Assert.Null(config);
        }
        
        [Fact]
        public void LoadConfig_ShouldReturnFalse_WhenFileIsInvalid()
        {
            // Arrange
            var configService = new ConfigService();
            string invalidFile = "invalid_config.json";
            
            // Create an invalid JSON file
            System.IO.File.WriteAllText(System.IO.Path.Combine("cfg", invalidFile), "invalid json");
            
            // Act
            var result = configService.LoadConfig(invalidFile, out var config);
            
            // Assert
            Assert.False(result);
            Assert.Null(config);
            
            // Clean up
            System.IO.File.Delete(System.IO.Path.Combine("cfg", invalidFile));
        }
    }
}