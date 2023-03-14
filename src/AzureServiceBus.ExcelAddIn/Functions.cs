using System.Threading.Tasks;
using Azure.Messaging.ServiceBus;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Microsoft.Extensions.DependencyInjection;

namespace AzureServiceBus.ExcelAddIn;

public static class MyFunctions
{
    [ExcelFunction(Description = "Add 2 numbers together, fancy!")]
    public static object AddNumbers(double x, double y)
    {
		return AsyncTaskUtil.RunTask(
			nameof(ExcelController.AddNumbers), 
			new object[] { x, y },
			async () => await CallAzureServiceBusAsync(x, y));
    }

    private static async Task<double> CallAzureServiceBusAsync(double x, double y)
    {
	    var controller = ContainerOperations.Container.GetRequiredService<ExcelController>();
	    return await controller.AddNumbers(x, y);
    }
}
