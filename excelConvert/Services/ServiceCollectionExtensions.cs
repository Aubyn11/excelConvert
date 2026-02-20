using Microsoft.Extensions.DependencyInjection;

namespace excelConvert.Services
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddExcelConvertServices(this IServiceCollection services)
        {
            // 注册服务
            services.AddSingleton<ExcelService>();
            services.AddSingleton<ConfigService>();
            services.AddSingleton<ConfigModelGenerator>();
            services.AddSingleton<IConfigFactory, ConfigFactory>();
            
            return services;
        }
    }
}