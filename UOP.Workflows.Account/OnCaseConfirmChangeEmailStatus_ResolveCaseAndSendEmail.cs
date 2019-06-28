using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;

namespace UOP.Workflows.Account
{
    public class OnCaseConfirmChangeEmailStatus_ResolveCaseAndSendEmail : BaseWorkflow
    {
        private const string Target = "Target";
        private const string Email = "emailaddress1";
        private const string NewEmail = "new_newemail";
        private const string AccointId = "accountid";
        private const string AccountName = "name";
        private const string ChangeEmailStatus = "new_changeemailstatus";
        private const string SubStatusReason = "new_substatusreason";
        private Guid _userId;
        private Guid _emailId;
        private Guid _accountId;

        public override void ExecuteWorkflowActivity(IExtendedExecutionContext context)
        {
            // Depth check to prevent infinite loop
            if (context.Depth > 1)
            {
                return;
            }

            // The InputParameters collection contains all the data passed in the message request
            if (context.InputParameters.Contains(Target) &&
                context.InputParameters[Target] is Entity)
            {
                // Obtain the target Case entity from the input parameters
                Entity target = (Entity)context.InputParameters[Target];
                try
                {
                    if (context.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase) &&
                        target.Attributes.Contains(ChangeEmailStatus) && target.Attributes[ChangeEmailStatus] != null)
                    {
                        context.Trace($"Entered {nameof(OnCaseConfirmChangeEmailStatus_ResolveCaseAndSendEmail)}.Execute()");
                        context.ClearPluginTraceLog(nameof(OnCaseConfirmChangeEmailStatus_ResolveCaseAndSendEmail));

                        if (context.PreImageEntity == null || context.PostImageEntity == null)
                        {
                            throw new ArgumentNullException(context.PreImageEntity == null ? 
                                nameof(context.PreImageEntity) : nameof(context.PostImageEntity));
                        }

                        int newEmailStatus = context.PostImageEntity.GetAttributeValue<OptionSetValue>(ChangeEmailStatus).Value;
                        int confirmedOptionSetValue = new OptionSetValue(100000002).Value; //Confirmed

                        if (newEmailStatus != confirmedOptionSetValue)
                        {
                            return;
                        }

                        // Retrieve the account related to the target case
                        _accountId = context.PostImageEntity.GetAttributeValue<EntityReference>("customerid").Id;
                        if (_accountId == null)
                        {
                            context.Trace($"Unable to retrieve the account. Aborting workflow execution.");
                            return;
                        }

                        // Retrieve all the active cases with title "Email Change Request" for the current account
                        var allCases = RetrieveAllEmailChangeRequestCases(context);
                        if (!allCases.Any())
                        {
                            context.Trace($"No active Email Change Request cases found for account with id {_accountId}");
                            return;
                        }

                        context.Trace($"Retrieved {allCases.Count} active Email Change Request cases for account id {_accountId}");

                        // Retrieve the latest active "Email Change Request" case for the current account and resolve it.
                        Incident latestCase = allCases.First();
                        ResolveCase(context, latestCase);

                        Entity account = context.OrganizationService.Retrieve(PluginXrm.Account.EntityLogicalName, _accountId, 
                            new ColumnSet(AccointId, AccountName, Email));
                        account.Attributes[Email] = latestCase.GetAttributeValue<string>(NewEmail);
                        context.Trace($"Updated account {account.GetAttributeValue<string>(AccountName)}'s email to " +
                                      $"{latestCase.GetAttributeValue<string>(NewEmail)}");
                        context.OrganizationService.Update(account);

                        //Retrieve all other active "Email Change Request" cases (if any) for the current account that need to be closed as duplicate.
                        var casesToClose = allCases.Skip(1).ToList();
                        if (!casesToClose.Any())
                        {
                            context.Trace($"No other active cases for account with id {_accountId}");
                        }
                        else
                        {
                            // Close all other cases setting their built-in status and status reason to Cancelled
                            foreach (var currentCase in casesToClose)
                            {
                                CloseCase(context, currentCase);
                            }

                            context.Trace($"Closed {casesToClose.Count} cases with status reason Cancelled for account with id {_accountId}");
                        }

                        // Send Email to the account
                        bool mailSent = SendEmail(context, account, latestCase);
                        context.Trace(mailSent
                            ? $"Sent the email successfully changed message to user {account.GetAttributeValue<string>(AccountName)} with id {_accountId}."
                            : $"Email was not send successfully to user with id {_accountId}");
                    }
                }
                catch (Exception ex)
                {
                    context.Trace($"An error occurred in {nameof(OnCaseConfirmChangeEmailStatus_ResolveCaseAndSendEmail)}. " +
                        $"Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    context.Trace($"Exiting {nameof(OnCaseConfirmChangeEmailStatus_ResolveCaseAndSendEmail)}.Execute()");
                }
            }
        }

        private List<Incident> RetrieveAllEmailChangeRequestCases(IExtendedExecutionContext context)
        {
            using (var crmContext = new ServiceContext(context.OrganizationService) { MergeOption = MergeOption.NoTracking })
            {
                context.Trace($"Attempting to retrieve all active Email Change Request cases for account with id {_accountId}");
                var allCases = (from incidents in crmContext.IncidentSet
                                join account in crmContext.AccountSet on incidents.CustomerId.Id equals account.Id
                                where incidents.CustomerId.Id == _accountId &&
                                      incidents.SubjectId.Id.Equals(Guid.Parse(EmailChangeRequestSubjectId)) &&
                                      incidents.StateCode == IncidentState.Active
                                orderby incidents.CreatedOn descending
                                select incidents)
                                .ToList();

                return allCases;
            }
        }

        private static void CloseCase(IExtendedExecutionContext context, Incident currentCase)
        {
            currentCase[ChangeEmailStatus] = new OptionSetValue(100000003); //Declined
            currentCase[SubStatusReason] = new OptionSetValue(100000001); //Cancelled - Duplicate
            context.OrganizationService.Update(currentCase);

            SetStateRequest setStateRequest = new SetStateRequest
            {
                EntityMoniker = new EntityReference(Incident.EntityLogicalName, currentCase.Id),
                State = new OptionSetValue((int)IncidentState.Cancelled),
                Status = new OptionSetValue(6),
            };

            context.OrganizationService.Execute(setStateRequest);
            context.TracingService.Trace($"Set case status to Cancelled for Case with ID {currentCase.Id}");
        }

        private void ResolveCase(IExtendedExecutionContext context, Incident latestCase)
        {
            // Change the case built-in status and status reason to Resolved and Problem solved respectively on the latest case created.
            // Change the Change Email Status to Approved on the latest case created. 
            latestCase[ChangeEmailStatus] = new OptionSetValue(100000004); //Approved
            latestCase[SubStatusReason] = new OptionSetValue(100000002); //Approved - Confirmed
            context.OrganizationService.Update(latestCase);

            // Create Incident Resolution
            var incidentResolutionResolved = new IncidentResolution
            {
                Subject = "Case Resolved",
                IncidentId = new EntityReference(Incident.EntityLogicalName, latestCase.Id),
                ActualEnd = DateTime.Now
            };

            // Close Incident
            var closeIncidentRequestResolved = new CloseIncidentRequest
            {
                IncidentResolution = incidentResolutionResolved,
                Status = new OptionSetValue(5)
            };

            context.OrganizationService.Execute(closeIncidentRequestResolved);
            context.Trace($"Set case status to Approved for Case with ID {latestCase.Id}");
        }

        private bool SendEmail(IExtendedExecutionContext context, Entity account, Incident latestCase)
        {
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
                Subject = "Email successfully changed",
                Description = $"Hello, {account.GetAttributeValue<string>(AccountName)}.{Environment.NewLine}" +
                              $"We are writing you to inform you that your primary email was successfully changed " +
                              $"to {account.GetAttributeValue<string>(Email)}.{Environment.NewLine}" +
                              $"Thank you for using our services!",
                DirectionCode = true,
                RegardingObjectId = new EntityReference(Incident.EntityLogicalName, latestCase.Id)
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
    }
}
