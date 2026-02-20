using Microsoft.Extensions.DependencyInjection;

namespace excelConvert.Services
{
    public static class ServiceProvider
    {
        private static IServiceProvider _instance;
        
        public static void Initialize()
        {
            var services = new ServiceCollection();
            services.AddExcelConvertServices();
            
            _instance = services.BuildServiceProvider();
        }
        
        public static T GetService<T>() where T : class
        {
            return _instance.GetService<T>();
        }
    }
}