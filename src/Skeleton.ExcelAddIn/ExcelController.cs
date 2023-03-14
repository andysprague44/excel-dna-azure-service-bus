using System;
using Microsoft.Extensions.Logging;

namespace Skeleton.ExcelAddIn;

internal class ExcelController
{
	private readonly ILogger _logger;
	private readonly AppSettings _settings;

	public ExcelController(
		AppSettings settings,
		ILogger logger)
	{
		_logger = logger ?? throw new ArgumentNullException(nameof(logger));
		_settings = settings ?? throw new ArgumentNullException(nameof(settings));
	}

	public void DoSomething()
	{
		_logger.LogInformation("Hello from ExcelController (app version = {Version})!", _settings.Version);
	}
}
