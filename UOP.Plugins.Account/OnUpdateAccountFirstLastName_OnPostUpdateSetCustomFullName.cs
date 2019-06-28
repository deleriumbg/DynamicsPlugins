using System;
using System.Linq;
using Microsoft.Xrm.Sdk;


namespace UOP.Plugins.Account
{
    public class OnUpdateAccountFirstLastName_OnPostUpdateSetCustomFullName : BasePlugin
    {
        private const string PreImageAlias = "Pre";
        private const string PostImageAlias = "Post";
        private const string Target = "Target"; 
        private const string FirstName = "new_firstname";
        private const string LastName = "new_lastname";
        private const string FullName = "new_customfullname";

        public override void Execute(ILocalPluginContext localContext)
        {
            // Depth check to prevent infinite loop
            if (localContext.PluginExecutionContext.Depth > 1)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request
            if (localContext.PluginExecutionContext.InputParameters.Contains(Target) && 
                localContext.PluginExecutionContext.InputParameters[Target] is Entity)
            {
                // Obtain the target entity from the input parameters
                Entity target = (Entity)localContext.PluginExecutionContext.InputParameters[Target];
                try
                {
                    // Check for update event
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnUpdateAccountFirstLastName_OnPostUpdateSetCustomFullName));
                        localContext.ClearPluginTraceLog(nameof(OnUpdateAccountFirstLastName_OnPostUpdateSetCustomFullName));
                        localContext.Trace($"Attempting to retrieve account data...");

                        // Obtain the pre image and post image entities
                        Entity preImageEntity = (localContext.PluginExecutionContext.PreEntityImages != null && 
                            localContext.PluginExecutionContext.PreEntityImages.Contains(PreImageAlias))
                            ? localContext.PluginExecutionContext.PreEntityImages[PreImageAlias]
                            : null;
                        Entity postImageEntity = (localContext.PluginExecutionContext.PostEntityImages != null && 
                            localContext.PluginExecutionContext.PostEntityImages.Contains(PostImageAlias))
                            ? localContext.PluginExecutionContext.PostEntityImages[PostImageAlias]
                            : null;

                        string previousFirstName = preImageEntity?.GetAttributeValue<string>(FirstName);
                        string newFirstName = postImageEntity?.GetAttributeValue<string>(FirstName);

                        string previousLastName = preImageEntity?.GetAttributeValue<string>(LastName);
                        string newLastName = postImageEntity?.GetAttributeValue<string>(LastName);

                        if (newFirstName == null && newLastName == null)
                        {
                            localContext.Trace($"Both First Name and Last Name fields cannot be empty. Account entity was not updated.");
                            return;
                        }

                        if (previousFirstName == newFirstName && previousLastName == newLastName)
                        {
                            localContext.Trace($"Newly entered First Name and Last Name are the same as the old ones. Account entity was not updated.");
                            return;
                        }

                        localContext.Trace($"First Name updated from {previousFirstName ?? "Empty field"} to {newFirstName}");
                        localContext.Trace($"Last Name updated from {previousLastName ?? "Empty field"} to {newLastName}");

                        // Join the first name and last name (ignoring nulls) 
                        string previousFullName = preImageEntity?.GetAttributeValue<string>(FullName);
                        string newFullName = string.Join(" ", new[] { newFirstName, newLastName }.Where(n => !string.IsNullOrWhiteSpace(n)));

                        // Create new instance of account for update
                        var retrievedAccount = new Entity(PluginXrm.Account.EntityLogicalName, target.Id);
                        var account = new Entity(PluginXrm.Account.EntityLogicalName)
                        {
                            Id = retrievedAccount.Id,
                            [FullName] = newFullName,
                            ["name"] = newFullName,
                        };

                        localContext.Trace($"Full Name updated successfully from {previousFullName ?? "Empty field"} to {newFullName}");
                        localContext.OrganizationService.Update(account);
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnUpdateAccountFirstLastName_OnPostUpdateSetCustomFullName)}. " +
                        $"Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    localContext.Trace("Exiting {0}.Execute()", nameof(OnUpdateAccountFirstLastName_OnPostUpdateSetCustomFullName));
                }
            }
        }
    }
}
