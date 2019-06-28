using System;
using System.Activities;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

namespace UOP.Workflows.Account
{
    public abstract class BaseWorkflow : CodeActivity
    {
        public const string EmailChangeRequestSubjectId = "1F5C140F-DA8D-E911-A97D-000D3A26C11D";

        protected override void Execute(CodeActivityContext activityContext)
        {
            if (activityContext == null)
            {
                throw new ArgumentNullException(nameof(activityContext));
            }

            // Get local workflow context
            var context = new BaseWorkflowActivityContext(activityContext, this);

            try
            {
                // Invoke the custom implementation 
                ExecuteWorkflowActivity(context);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                context.Trace($"Exception: {e}");

                // Handle the exception.
                throw;
            }
        }

        /// <summary>
        /// Execution method for the workflow activity
        /// </summary>
        /// <param name="context">Context for the current plug-in.</param>
        public abstract void ExecuteWorkflowActivity(IExtendedExecutionContext context);
    }
}
