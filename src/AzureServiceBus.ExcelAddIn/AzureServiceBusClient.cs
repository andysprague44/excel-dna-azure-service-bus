using System;
using System.Text.Json;
using System.Threading.Tasks;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;
using Microsoft.Extensions.Logging;

namespace AzureServiceBus.ExcelAddIn;

/// <summary>
/// Generic interface to do some work.
/// In this example it's implemented by an AzureServiceBusClient that offloads the work
/// </summary>
public interface ITaskProcessor
{
	public Task<TR?> ExecuteAsync<T, TR>(string functionName, T request);
}

internal class AzureServiceBusClient : ITaskProcessor
{
	private readonly string _azureServiceBusConnectionString;
	private readonly string _mqPrefix;
	private readonly TimeSpan _maxTimeOut;
	private readonly ILogger _logger;
	
	internal AzureServiceBusClient(
		string connectionString,
		AppSettings settings,
		ILogger logger)
	{
		_azureServiceBusConnectionString = connectionString;
		_mqPrefix = settings.MQPrefix ?? throw new ArgumentException(nameof(settings.MQPrefix));
		_maxTimeOut = settings.MaxTimeout ?? TimeSpan.FromSeconds(30);
		_logger = logger ?? throw new ArgumentNullException(nameof(logger));
	}

	public async Task<TR?> ExecuteAsync<T, TR>(string functionName, T request)
	{
		//Set up queues if required
		var requestQueue = _mqPrefix + $"{functionName.ToLower()}.inq";
		await CreateQueueIfNotExists(requestQueue);

		var responseQueue = _mqPrefix + $"{functionName.ToLower()}.outq";
		await CreateQueueIfNotExists(responseQueue);

		await using var client = new ServiceBusClient(_azureServiceBusConnectionString);

		//Send
		var jsonRequest = JsonSerializer.Serialize(request);
		var message = new ServiceBusMessage(jsonRequest)
		{
			ReplyTo = responseQueue,
		};


		try
		{

			await using var sender = client.CreateSender(requestQueue);
			await sender.SendMessageAsync(message);
		}
		catch (Exception exRequest)
		{
			_logger.LogCritical(exRequest, "Failed to send request: {ex}", exRequest.Message);
			throw;
		}

		//In a real application you would have the message processed by another service
		await MockOutABackgroundServicePollingTheQueue(client);

		//Receive
		try
		{
			await using var receiver = client.CreateReceiver(message.ReplyTo);
			var receivedMessage = await receiver.ReceiveMessageAsync(_maxTimeOut);
			var responseJson = receivedMessage.Body.ToString();
			var response = JsonSerializer.Deserialize<TR>(responseJson);

			return response;
		}
		catch (Exception exResponse)
		{
			_logger.LogCritical(exResponse, "Failed to recieve response: {ex}", exResponse.Message);
			throw;
		}
	}

	private async Task CreateQueueIfNotExists(string queue)
	{
		try
		{
			var administrationClient = new ServiceBusAdministrationClient(_azureServiceBusConnectionString);
			var queueExists = await administrationClient.QueueExistsAsync(queue);
			if (!queueExists)
			{
				await administrationClient.CreateQueueAsync(queue);
			}
		}
		catch (Exception ex)
		{
			_logger.LogCritical(ex, "Failed to create queue {QueueName}: {ex}", queue, ex.Message);
			throw;
		}
	}

	private async Task MockOutABackgroundServicePollingTheQueue(ServiceBusClient client)
	{
		//this method wouldn't be part of the excel add in, but a background service processing messages from the queue
		//adding here for illustration only

		//Receive request
		try
		{
			await using var receiver = client.CreateReceiver(_mqPrefix + "addnumbers.inq");
			var receivedMessage = await receiver.ReceiveMessageAsync(_maxTimeOut);
			var responseJson = receivedMessage.Body.ToString();

			// The background service knows the format of the request, based on the queue name
			// AddNumbersRequest/AddNumbersResponse is a shared contract
			var request = JsonSerializer.Deserialize<ExcelController.AddNumbersRequest>(responseJson)!;

			//Process
			var response = new ExcelController.AddNumbersResponse
			{
				Z = request.X + request.Y,
			};

			//Send response
			await using var sender = client.CreateSender(receivedMessage.ReplyTo);
			var jsonResponse = JsonSerializer.Serialize(response);
			var message = new ServiceBusMessage(jsonResponse);
			await sender.SendMessageAsync(message);
		}
		catch (Exception ex)
		{
			_logger.LogCritical(ex, "Failed to process message: {ex}", ex.Message);
			throw;
		}
	}
}
