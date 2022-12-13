using Microsoft.Extensions.DependencyInjection;
using OEM.Core;
using OEM.Npoi;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Extensions.DependencyInjection
{
    public static class OEMServiceCollectionExtensions
    {
        public static void AddExcelToObjectNpoiService(this IServiceCollection services)
        {
            services.AddSingleton(typeof(IExcelFactory), typeof(NpoiExcelFactory));
        }
    }
}
