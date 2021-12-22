using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging.EventLog;

namespace BulkRegistrationProcesses
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .UseWindowsService()
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddHostedService<Worker>()
                      .Configure<EventLogSettings>(config =>
                      {
                          config.LogName = "BulkRegistrationProcesses";
                          config.SourceName = "BulkRegistrationProcessesSource";
                      });
                } );
    }
}