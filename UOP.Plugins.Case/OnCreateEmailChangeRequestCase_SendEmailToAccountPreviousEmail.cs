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
        private const string CustomerId = "customerid";
        private Guid _emailId;
        private Guid _accountId;
        private Guid _userId;

        public override void Execute(ILocalPluginContext localContext)
        {
            // Depth check to prevent infinite loop
            if (localContext.PluginExecutionContext.Depth > 2)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request
            if (localContext.PluginExecutionContext.InputParameters.Contains(Target) && 
                localContext.PluginExecutionContext.InputParameters[Target] is Entity)
            {
                // Obtain the target Case entity from the input parameters
                Entity target = (Entity)localContext.PluginExecutionContext.InputParameters[Target];
                try
                {
                    // Check for create event for case with subject Email Change Request
                    if (localContext.PluginExecutionContext.MessageName.Equals("create", StringComparison.InvariantCultureIgnoreCase) && 
                        target["subjectid"].Equals(new EntityReference(Subject.EntityLogicalName, Guid.Parse(EmailChangeRequestSubjectId))))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail));
                        localContext.ClearPluginTraceLog(nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail));
                        localContext.Trace($"Attempting to retrieve case with title Email Change Request...");

                        // Get a system user to send the email (From: field)
                        _userId = localContext.PluginExecutionContext.UserId;

                        // Check if case customer is account or contact
                        var customerType = ((EntityReference)target[CustomerId]).LogicalName;
                        _accountId = target.GetAttributeValue<EntityReference>(CustomerId).Id;
                        localContext.Trace($"Retrieved {customerType} with id {_accountId} related to the case");

                        // Get the email recipient user
                        // Retrieving all fields (ColumnSet(true)), because field names are different for account and contact entity
                        Entity recipient = localContext.OrganizationService.Retrieve(customerType, _accountId, new ColumnSet(true));
                        if (recipient == null)
                        {
                            localContext.Trace($"Unable to retrieve the account related to the target case with id {target.Id}. " +
                                               $"Aborting plugin execution.");
                            return;
                        }

                        // Create the 'From:' activity party for the email
                        ActivityParty fromParty = new ActivityParty
                        {
                            PartyId = new EntityReference(SystemUser.EntityLogicalName, _userId)
                        };

                        // Create the 'To:' activity party for the email
                        ActivityParty toParty = new ActivityParty
                        {
                            PartyId = new EntityReference(customerType, _accountId),
                            AddressUsed = target.GetAttributeValue<string>("new_previousemail")
                        };

                        localContext.Trace("Created To and From activity parties.");

                        // Field names are different for account is "name" and for contact is "fullname"
                        var accountName = customerType == "account" ? "name" : "fullname";

                        // Create an email message entity
                        Email email = new Email
                        {
                            To = new ActivityParty[] { toParty },
                            From = new ActivityParty[] { fromParty },
                            Subject = "Email change request confirmation",
                            Description = $"Hello, {recipient.GetAttributeValue<string>(accountName)}.{Environment.NewLine}" +
                                          $"We received your request to change your primary email address from " +
                                          $"{target.GetAttributeValue<string>("new_previousemail")} to " +
                                          $"{target.GetAttributeValue<string>("new_newemail")}.{Environment.NewLine}" +
                                          $"Please CLICK HERE to confirm your request",
                            DirectionCode = true,
                            RegardingObjectId = new EntityReference(Incident.EntityLogicalName, target.Id)
                        };

                        _emailId = localContext.OrganizationService.Create(email);
                        localContext.Trace($"Created {email.Subject} email with description: {email.Description}.");

                        SendEmailRequest sendEmailRequest = new SendEmailRequest
                        {
                            EmailId = _emailId,
                            TrackingToken = string.Empty,
                            IssueSend = true
                        };

                        SendEmailResponse sendEmailResponse = (SendEmailResponse)localContext.OrganizationService.Execute(sendEmailRequest);
                        localContext.Trace($"Sent the confirm email change request message to user with id {target.Id}.");
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnCreateEmailChangeRequestCase_SendEmailToAccountPreviousEmail)}. " +
                        $"Exception details: {ex.Message}");
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
