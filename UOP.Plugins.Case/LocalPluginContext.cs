using System;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;

namespace UOP.Plugins.Case
{
    public class LocalPluginContext : ILocalPluginContext
    {
        private readonly IServiceProvider _serviceProvider;
        private IPluginExecutionContext _pluginExecutionContext;
        private ITracingService _tracingService;
        private IOrganizationServiceFactory _organizationServiceFactory;
        private IOrganizationService _organizationService;

        public IOrganizationService OrganizationService =>
            _organizationService ??
            (_organizationService = OrganizationServiceFactory.CreateOrganizationService(PluginExecutionContext.UserId));

        public IPluginExecutionContext PluginExecutionContext =>
            _pluginExecutionContext ??
            (_pluginExecutionContext = (IPluginExecutionContext)_serviceProvider.GetService(typeof(IPluginExecutionContext)));

        public ITracingService TracingService => _tracingService ?? (_tracingService = (ITracingService)_serviceProvider.GetService(typeof(ITracingService)));

        private IOrganizationServiceFactory OrganizationServiceFactory =>
            _organizationServiceFactory ??
            (_organizationServiceFactory = (IOrganizationServiceFactory)_serviceProvider.GetService(typeof(IOrganizationServiceFactory)));

        public LocalPluginContext(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
        }

        public void Trace(string message, params object[] o)
        {
            if (PluginExecutionContext == null)
            {
                SafeTrace(message, o);
            }
            else
            {
                SafeTrace(
                    "{0}, Correlation Id: {1}, Initiating User: {2}",
                    string.Format(message, o),
                    PluginExecutionContext.CorrelationId,
                    PluginExecutionContext.InitiatingUserId);
            }
        }

        public void Trace(FaultException<OrganizationServiceFault> exception)
        {
            // Trace the first message using the embedded Trace to get the Correlation Id and User Id out.
            Trace("Exception: {0}", exception.Message);

            // From here on use the tracing service trace
            SafeTrace(exception.StackTrace);

            if (exception.Detail != null)
            {
                SafeTrace("Error Code: {0}", exception.Detail.ErrorCode);
                SafeTrace("Detail Message: {0}", exception.Detail.Message);
                if (!string.IsNullOrEmpty(exception.Detail.TraceText))
                {
                    SafeTrace("Trace: ");
                    SafeTrace(exception.Detail.TraceText);
                }

                foreach (var item in exception.Detail.ErrorDetails)
                {
                    SafeTrace("Error Details: ");
                    SafeTrace(item.Key);
                    SafeTrace(item.Value.ToString());
                }

                if (exception.Detail.InnerFault != null)
                {
                    Trace(new FaultException<OrganizationServiceFault>(exception.Detail.InnerFault));
                }
            }
        }

        private void SafeTrace(string message, params object[] o)
        {
            if (string.IsNullOrWhiteSpace(message) || TracingService == null)
            {
                return;
            }
            TracingService.Trace(message, o);
        }

        public void ClearPluginTraceLog(string className)
        {
            using (var crmContext = new ServiceContext(OrganizationService) { MergeOption = MergeOption.NoTracking })
            {
                var pluginTraceLogs = (from traceLogs in crmContext.PluginTraceLogSet
                                       where traceLogs.TypeName == $"{this.GetType().Namespace}.{className}"
                                       select traceLogs)
                                       .ToList();

                this.Trace($"Deleting {pluginTraceLogs.Count} plugin trace logs of type {this.GetType().Namespace}.{className}");
                foreach (var traceLog in pluginTraceLogs)
                {
                    this.OrganizationService.Delete(PluginTraceLog.EntityLogicalName, traceLog.Id);
                }
            }
        }
    }
}
