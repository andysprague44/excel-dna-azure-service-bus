using ExcelDna.Integration;
using Microsoft.Extensions.DependencyInjection;

namespace Skeleton.ExcelAddIn;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string HelloDna(string name)
    {
		//To prove we can get to the controller
		var controller = ContainerOperations.Container.GetRequiredService<ExcelController>();
		controller.DoSomething();

        return "Hello " + name;
    }
}
