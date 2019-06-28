using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;
using System;


namespace UOP.Workflows.Account
{
    public class OnUpdateAccountEmail_CreateCaseAndSendEmail : BaseWorkflow
    {
        private const string Target = "Target";
        private const string AccountId = "accountid";
        private const string AccountName = "name";
        private const string Email = "emailaddress1";
        private const string PreviousEmail = "new_previousemail";
        private const string NewEmail = "new_newemail";
        private Guid _emailId;
        private Guid _accountId;
        private Guid _userId;
        private Guid _caseId;

        public override void ExecuteWorkflowActivity(IExtendedExecutionContext context)
        {
            // Depth check to prevent infinite loop
            if (context.Depth > 1)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains(Target) &&
                context.InputParameters[Target] is Entity)
            {
                // Obtain the target Account entity from the input parameters.
                Entity target = (Entity)context.InputParameters[Target];
                try
                {
                    if (context.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase))
                    {
                        context.Trace($"Entered {nameof(OnUpdateAccountEmail_CreateCaseAndSendEmail)}.Execute()");
                        context.ClearPluginTraceLog(nameof(OnUpdateAccountEmail_CreateCaseAndSendEmail));
                        context.Trace($"Attempting to retrieve account email...");

                        if (context.PreImageEntity == null || context.PostImageEntity == null)
                        {
                            throw new ArgumentNullException(context.PreImageEntity == null ? 
                                nameof(context.PreImageEntity) : nameof(context.PostImageEntity));
                        }

                        string previousEmail = context.PreImageEntity.GetAttributeValue<string>(Email);
                        string newEmail = context.PostImageEntity.GetAttributeValue<string>(Email);

                        if (previousEmail == newEmail)
                        {
                            context.Trace($"Newly entered Email is the same as the old one. Aborting workflow execution.");
                            return;
                        }

                        // Create new instance of account for update
                        var retrievedAccount = new Entity(PluginXrm.Account.EntityLogicalName, target.Id);
                        var accountInstance = new Entity(PluginXrm.Account.EntityLogicalName)
                        {
                            Id = retrievedAccount.Id,
                            [Email] = previousEmail
                        };

                        context.Trace($"Email field value set to the previous one: {previousEmail}");
                        context.OrganizationService.Update(accountInstance);

                        _accountId = target.Id;
                        Entity account = context.OrganizationService.Retrieve(PluginXrm.Account.EntityLogicalName, _accountId, new ColumnSet(AccountId, AccountName, Email));
                        if (account == null)
                        {
                            context.Trace($"Unable to retrieve the account with id {_accountId}. " +
                                          $"Aborting workflow execution.");
                        }

                        // Create Case
                        Incident caseRecord = CreateCase(context, previousEmail, newEmail);
                        _caseId = context.OrganizationService.Create(caseRecord);
                        context.Trace($"Case with id {_caseId} created ");

                        // Send Email
                        bool mailSent = SendEmail(context, account, caseRecord);
                        context.Trace(mailSent
                            ? $"Sent the confirm email change request message to user with id {_accountId}."
                            : $"Email was not send successfully to user with id {_accountId}");
                    }
                }
                catch (Exception ex)
                {
                    context.Trace($"An error occurred in {nameof(OnUpdateAccountEmail_CreateCaseAndSendEmail)}. " +
                        $"Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    context.Trace($"Exiting {nameof(OnUpdateAccountEmail_CreateCaseAndSendEmail)}.Execute()");
                }
            }
        }

        private bool SendEmail(IExtendedExecutionContext context, Entity account, Incident caseRecord)
        {
            context.Trace($"Attempting to retrieve case with title Email Change Request...");

            // Get a system user to send the email (From: field)
            _userId = context.UserId;

            // Create the 'From:' activity party for the email
            ActivityParty fromParty = new ActivityParty
            {
                PartyId = new EntityReference(SystemUser.EntityLogicalName, _userId)
            };

            // Create the 'To:' activity party for the email
            ActivityParty toParty = new ActivityParty
            {
                PartyId = new EntityReference(PluginXrm.Account.EntityLogicalName, _accountId),
                AddressUsed = account.GetAttributeValue<string>(Email)
            };

            context.Trace("Created To and From activity parties.");

            // Create an email message entity
            Email email = new Email
            {
                To = new ActivityParty[] {toParty},
                From = new ActivityParty[] {fromParty},
                Subject = "Email change request confirmation",
                Description = $"Hello, {account.GetAttributeValue<string>(AccountName)}.{Environment.NewLine}" +
                              $"We received your request to change your primary email address from " +
                              $"{caseRecord.GetAttributeValue<string>(PreviousEmail)} to " +
                              $"{caseRecord.GetAttributeValue<string>(NewEmail)}.{Environment.NewLine}" +
                              @"Please <a href=""https://www.uopeople.edu/"">Click Here</a> to confirm your request",
                DirectionCode = true,
                RegardingObjectId = new EntityReference(Incident.EntityLogicalName, _caseId)
            };

            _emailId = context.OrganizationService.Create(email);
            context.Trace($"Created {email.Subject} email with description: {email.Description}.");

            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = _emailId,
                TrackingToken = string.Empty,
                IssueSend = true
            };

            SendEmailResponse sendEmailResponse = (SendEmailResponse)context.OrganizationService.Execute(sendEmailRequest);
            return sendEmailResponse != null;
        }

        private Incident CreateCase(IExtendedExecutionContext context, string previousEmail, string newEmail)
        {
            Incident incident = new Incident();
            context.Trace($"Creating new case...");

            incident["title"] = "Email Change Request";
            context.Trace($"Case Title set to Email Change Request");

            incident[PreviousEmail] = previousEmail;
            context.Trace($"Case Previous Email set to {previousEmail}");

            incident[NewEmail] = newEmail;
            context.Trace($"Case New Email set to {newEmail}");

            incident["subjectid"] = new EntityReference(Subject.EntityLogicalName, Guid.Parse(EmailChangeRequestSubjectId));
            context.Trace($"Case Subject set to Email Change Request");

            incident.Attributes.Add("customerid", new EntityReference(PluginXrm.Account.EntityLogicalName, _accountId));
            context.Trace($"Case connected with account id {_accountId}");

            return incident;
        }
    }
}
