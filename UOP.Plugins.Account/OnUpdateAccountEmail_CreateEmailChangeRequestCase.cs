using System;
using Microsoft.Xrm.Sdk;


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
            //Depth check to prevent infinite loop
            if (localContext.PluginExecutionContext.Depth > 1)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request.
            if (localContext.PluginExecutionContext.InputParameters.Contains(Target) && localContext.PluginExecutionContext.InputParameters[Target] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)localContext.PluginExecutionContext.InputParameters[Target];
                try
                {
                    //If it's update get the latest data 
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                        localContext.ClearPluginTraceLog(localContext, nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                        localContext.Trace($"Attempting to retrieve account email...");

                        //Obtain the pre image and post image entities
                        Entity preImageEntity =
                            (localContext.PluginExecutionContext.PreEntityImages != null && localContext.PluginExecutionContext.PreEntityImages.Contains(PreImageAlias))
                                ? localContext.PluginExecutionContext.PreEntityImages[PreImageAlias]
                                : null;
                        Entity postImageEntity =
                            (localContext.PluginExecutionContext.PostEntityImages != null && localContext.PluginExecutionContext.PostEntityImages.Contains(PostImageAlias))
                                ? localContext.PluginExecutionContext.PostEntityImages[PostImageAlias]
                                : null;

                        if (preImageEntity == null || !preImageEntity.Attributes.Contains(Email) ||
                            preImageEntity.GetAttributeValue<string>(Email) == null)
                        {
                            throw new InvalidPluginExecutionException(
                                $"An error occurred in {nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase)} plugin. The preImage fields are empty!");
                        }

                        if (postImageEntity == null || !postImageEntity.Attributes.Contains(Email) ||
                            postImageEntity.GetAttributeValue<string>(Email) == null)
                        {
                            throw new InvalidPluginExecutionException(
                                $"An error occurred in {nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase)} plugin. The new value for Email field is null.");
                        }

                        string previousEmail = preImageEntity.GetAttributeValue<string>(Email);
                        string newEmail = postImageEntity.GetAttributeValue<string>(Email);

                        if (previousEmail == newEmail)
                        {
                            localContext.Trace($"Newly entered Email is the same as the old one. Aborting plugin execution.");
                            return;
                        }

                        target[Email] = previousEmail;
                        localContext.Trace($"Email field value set to the previous one: {previousEmail}");
                        localContext.OrganizationService.Update(target);

                        //Create case
                        Entity caseRecord = CreateCase(localContext.TracingService, previousEmail, newEmail, target.Id);
                        localContext.OrganizationService.Create(caseRecord);
                        localContext.Trace($"Case with id {caseRecord.Id} created ");
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase)}. Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    localContext.Trace("Exiting {0}.Execute()", nameof(OnUpdateAccountEmail_CreateEmailChangeRequestCase));
                }
            }
        }

        private Entity CreateCase(ITracingService tracingService, string previousEmail, string newEmail, Guid accountId)
        {
            Entity caseRecord = new Entity("incident");
            tracingService.Trace($"Creating new case...");

            caseRecord["title"] = "Email Change Request";
            tracingService.Trace($"Case Title set to Email Change Request");

            caseRecord["new_previousemail"] = previousEmail;
            tracingService.Trace($"Case Previous Email set to {previousEmail}");

            caseRecord["new_newemail"] = newEmail;
            tracingService.Trace($"Case New Email set to {newEmail}");

            caseRecord["subjectid"] = new EntityReference("subject", Guid.Parse("1F5C140F-DA8D-E911-A97D-000D3A26C11D"));
            tracingService.Trace($"Case Subject set to Email Change Request");

            caseRecord.Attributes.Add("customerid", new EntityReference("account", accountId));
            tracingService.Trace($"Case connected with account id {accountId}");

            return caseRecord;
        }
    }
}
