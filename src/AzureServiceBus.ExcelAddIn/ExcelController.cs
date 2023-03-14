using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace AzureServiceBus.ExcelAddIn;

internal class ExcelController
{
	private readonly ITaskProcessor _processor;
	private readonly ILogger _logger;

	public ExcelController(
		ITaskProcessor processor,
		ILogger logger)
	{
		_processor = processor;
		_logger = logger ?? throw new ArgumentNullException(nameof(logger));
	}
	
	public async Task<double> AddNumbers(double x, double y)
	{
		_logger.LogInformation("Executing {Function} with x={X}, y={Y}",nameof(AddNumbers), x, y);

		var request = new AddNumbersRequest
		{
			X = x,
			Y = y,
		};
		var response = await _processor.ExecuteAsync<AddNumbersRequest, AddNumbersResponse>(nameof(AddNumbers), request);
		if (response == null)
			throw new Exception("Null response received");
		
		_logger.LogInformation("Completed executing {Function}. Got response={Response}",nameof(AddNumbers), response);
		
		return response.Z;
	}

	public class AddNumbersRequest
	{
		public double X { get; set; }
		public double Y { get; set; }

		public override string ToString()
		{
			return $"X={X}, Y={Y}";
		}
	}

	public class AddNumbersResponse
	{
		public double Z { get; set;  }
		
		public override string ToString()
		{
			return $"Z={Z}";
		}
	}
}
