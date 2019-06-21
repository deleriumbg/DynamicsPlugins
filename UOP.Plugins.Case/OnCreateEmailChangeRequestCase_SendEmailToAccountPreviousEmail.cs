using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;

namespace UOP.Plugins.Case
{
    public class OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail : BasePlugin
    {
        private const string Target = "Target";
        private Guid _emailId;
        private Guid _accountId;
        private Guid _userId;

        public override void Execute(ILocalPluginContext localContext)
        {
            //Depth check to prevent infinite loop
            if (localContext.PluginExecutionContext.Depth > 2)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request
            if (localContext.PluginExecutionContext.InputParameters.Contains(Target) && localContext.PluginExecutionContext.InputParameters[Target] is Entity)
            {
                // Obtain the target Case entity from the input parameters
                Entity target = (Entity)localContext.PluginExecutionContext.InputParameters[Target];
                try
                {
                    //Check for create event for case with title Email Change Request
                    if (localContext.PluginExecutionContext.MessageName.Equals("create", StringComparison.InvariantCultureIgnoreCase) && target["title"].Equals("Email Change Request"))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail));
                        localContext.ClearPluginTraceLog(localContext, nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail));
                        localContext.Trace($"Attempting to retrieve case with title Email Change Request...");

                        // Get a system user to send the email (From: field)
                        WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                        WhoAmIResponse systemUserResponse = (WhoAmIResponse)localContext.OrganizationService.Execute(systemUserRequest);
                        _userId = systemUserResponse.UserId;

                        //Get the email recipient user(Account entity)
                        EntityReference customer = (EntityReference)target.Attributes["customerid"];
                        Entity account = localContext.OrganizationService.Retrieve("account", customer.Id, new ColumnSet(true));
                        if (account == null)
                        {
                            localContext.Trace($"Unable to retrieve the account related to the target case with id {target.Id}. Aborting plugin execution.");
                            return;
                        }
                        _accountId = account.Id;

                        // Create the 'From:' activity party for the email
                        ActivityParty fromParty = new ActivityParty
                        {
                            PartyId = new EntityReference(SystemUser.EntityLogicalName, _userId)
                        };

                        // Create the 'To:' activity party for the email
                        ActivityParty toParty = new ActivityParty
                        {
                            PartyId = new EntityReference(Account.EntityLogicalName, _accountId),
                            AddressUsed = target.GetAttributeValue<string>("new_previousemail")
                        };

                        localContext.Trace("Created To and From activity parties.");

                        // Create an email message entity
                        Email email = new Email
                        {
                            To = new ActivityParty[] { toParty },
                            From = new ActivityParty[] { fromParty },
                            Subject = "Email change request confirmation",
                            Description = $"Hello, {account.GetAttributeValue<string>("name")}.{Environment.NewLine}" +
                                          $"We received your request to change your primary email address from {target.GetAttributeValue<string>("new_previousemail")} " +
                                          $"to {target.GetAttributeValue<string>("new_newemail")}.{Environment.NewLine}" +
                                          $"Please CLICK HERE to confirm your request",
                            DirectionCode = true,
                            RegardingObjectId = new EntityReference("incident", target.Id)
                        };

                        _emailId = localContext.OrganizationService.Create(email);
                        localContext.Trace($"Created {email.Subject} email with description: {email.Description}.");

                        SendEmailRequest sendEmailRequest = new SendEmailRequest
                        {
                            EmailId = _emailId,
                            TrackingToken = "",
                            IssueSend = true
                        };

                        SendEmailResponse sendEmailResponse = (SendEmailResponse)localContext.OrganizationService.Execute(sendEmailRequest);
                        localContext.Trace($"Sent the confirm email change request message to user with id {target.Id}.");
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail)}. Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    localContext.Trace("Exiting {0}.Execute()", nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail));
                }
            }
        }
    }
}
