using Microsoft.Extensions.Configuration;
using Microsoft.VisualBasic.CompilerServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using Microsoft.Extensions.Configuration.FileExtensions;

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
