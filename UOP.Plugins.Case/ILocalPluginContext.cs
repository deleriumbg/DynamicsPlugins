using System.ServiceModel;
using Microsoft.Xrm.Sdk;

namespace UOP.Plugins.Case
{
    public interface ILocalPluginContext
    {
        IOrganizationService OrganizationService { get; }
        IPluginExecutionContext PluginExecutionContext { get; }
        ITracingService TracingService { get; }
        void Trace(string message, params object[] o);
        void Trace(FaultException<OrganizationServiceFault> exception);
        void ClearPluginTraceLog(string className);
    }
}
