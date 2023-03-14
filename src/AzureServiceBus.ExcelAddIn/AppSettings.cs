using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureServiceBus.ExcelAddIn
{
	internal class AppSettings
	{
		public string Version { get; set; } = "0.1";
		public string? MQPrefix { get; set; }
		public TimeSpan? MaxTimeout { get; set; }
	}
}
