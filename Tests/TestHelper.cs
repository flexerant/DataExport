using Microsoft.Extensions.Configuration;
using System.IO;

namespace Tests
{
    class TestHelpers
    {
        private static readonly IConfiguration _configuration;
        private static readonly TestSettings _testSettings = new TestSettings();

        static TestHelpers()
        {
            _configuration = new ConfigurationBuilder()
             .SetBasePath(Directory.GetCurrentDirectory())
             .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
             .Build();

            _configuration.GetSection("TestSettings").Bind(_testSettings);
        }

        public static TestSettings TestSettings  => _testSettings;
    }
}
