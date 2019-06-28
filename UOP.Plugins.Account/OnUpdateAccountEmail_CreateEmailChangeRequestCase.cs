using System;
using Microsoft.Xrm.Sdk;
using PluginXrm;


namespace UOP.Plugins.Account
{
    public class OnUpdateAccountEmail_CreateEmailChangeRequestCase : BasePlugin
    {
        private const string PreImageAlias = "Pre";
        private const string PostImageAlias = "Post";
        private const string Target = "Target";
        private const string Email = "emailaddress1";

        public override void Execute(ILocalPluginContext localContext)
        {
            // Depth check to prevent infinite loop
            if (localContext.PluginExecutionContext.Depth > 1)
            {
                return;
            }
            
            // The InputParameters collection contains all the data passed in the message request.
            if (localContext.PluginExecutionContext.InputParameters.Contains(Target) && 
                localContext.PluginExecutionContext.InputParameters[Target] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)localContext.PluginExecutionContext.InputParameters[Target];
                try
                {
                    // Check for update event
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                        localContext.ClearPluginTraceLog(nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                        localContext.Trace($"Attempting to retrieve account email...");

                        //Obtain the pre image and post image entities
                        Entity preImageEntity = (localContext.PluginExecutionContext.PreEntityImages != null && 
                            localContext.PluginExecutionContext.PreEntityImages.Contains(PreImageAlias))
                                ? localContext.PluginExecutionContext.PreEntityImages[PreImageAlias]
                                : null;
                        Entity postImageEntity = (localContext.PluginExecutionContext.PostEntityImages != null && 
                            localContext.PluginExecutionContext.PostEntityImages.Contains(PostImageAlias))
                                ? localContext.PluginExecutionContext.PostEntityImages[PostImageAlias]
                                : null;

                        string previousEmail = preImageEntity?.GetAttributeValue<string>(Email);
                        string newEmail = postImageEntity?.GetAttributeValue<string>(Email);

                        if (previousEmail == newEmail)
                        {
                            localContext.Trace($"Newly entered Email is the same as the old one. Aborting plugin execution.");
                            return;
                        }

                        // Create new instance of account for update
                        var retrievedAccount = new Entity(PluginXrm.Account.EntityLogicalName, target.Id);
                        var account = new Entity(PluginXrm.Account.EntityLogicalName)
                        {
                            Id = retrievedAccount.Id,
                            [Email] = previousEmail
                        };

                        localContext.OrganizationService.Update(account);
                        localContext.Trace($"Email field value set to the previous one: {previousEmail}");

                        // Create case
                        Incident caseRecord = CreateCase(localContext, previousEmail, newEmail, account.Id);
                        var caseId = localContext.OrganizationService.Create(caseRecord);
                        localContext.Trace($"Case with id {caseId} created");
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase)}. " +
                        $"Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    localContext.Trace("Exiting {0}.Execute()", nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                }
            }
        }

        private Incident CreateCase(ILocalPluginContext localContext, string previousEmail, string newEmail, Guid accountId)
        {
            Incident incident = new Incident();
            localContext.Trace($"Creating new case...");

            incident["title"] = "Email Change Request";
            localContext.Trace($"Case Title set to Email Change Request");

            incident["new_previousemail"] = previousEmail;
            localContext.Trace($"Case Previous Email set to {previousEmail}");

            incident["new_newemail"] = newEmail;
            localContext.Trace($"Case New Email set to {newEmail}");

            incident["subjectid"] = new EntityReference(Subject.EntityLogicalName, Guid.Parse(EmailChangeRequestSubjectId));
            localContext.Trace($"Case Subject set to Email Change Request");

            incident.Attributes.Add("customerid", new EntityReference(PluginXrm.Account.EntityLogicalName, accountId));
            localContext.Trace($"Case connected with account id {accountId}");

            return incident;
        }
    }
}
