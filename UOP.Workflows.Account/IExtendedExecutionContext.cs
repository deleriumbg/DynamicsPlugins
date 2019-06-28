using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

namespace UOP.Workflows.Account
{
    public interface IExtendedExecutionContext : ITracingService, IOrganizationService, IWorkflowContext
    {
        /// <summary>
        /// Fullname of the workflow activity
        /// </summary>
        string WorkflowTypeName { get; }

        /// <summary>
        /// Extends ActivityContext and provides additional functionality for CodeActivity
        /// </summary>
        CodeActivityContext ActivityContext { get; }

        /// <summary>
        /// Provides logging run-time trace information
        /// </summary>
        ITracingService TracingService { get; }

        /// <summary>
        /// <see cref="IOrganizationService"/> using the user from the context
        /// </summary>
        IOrganizationService OrganizationService { get; }

        /// <summary>
        /// <see cref="PreImageEntity"/> Pre Image Entity
        /// </summary>
        Entity PreImageEntity { get; }

        /// <summary>
        /// <see cref="PostImageEntity"/> Post Image Entity
        /// </summary>
        Entity PostImageEntity { get; }

        /// <summary>
        /// <see cref="className"/> Delete all plugin trace logs with given name
        /// </summary>
        void ClearPluginTraceLog(string className);
    }
}
