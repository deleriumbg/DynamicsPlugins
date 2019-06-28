using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

namespace UOP.Plugins.Case
{
    public abstract class BasePlugin : IPlugin
    {
        private readonly string _className;
        public const string EmailChangeRequestSubjectId = "1F5C140F-DA8D-E911-A97D-000D3A26C11D";

        protected BasePlugin()
        {
            _className = GetType().Name;
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            // Construct the Local plug-in context.
            var localContext = new LocalPluginContext(serviceProvider);

            try
            {
                Execute(localContext);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                // Trace the exception before bubbling so that we ensure everything we need hits the log
                localContext.Trace(e);

                // Bubble the exception
                throw;
            }
        }

        public abstract void Execute(ILocalPluginContext localContext);
    }
}
