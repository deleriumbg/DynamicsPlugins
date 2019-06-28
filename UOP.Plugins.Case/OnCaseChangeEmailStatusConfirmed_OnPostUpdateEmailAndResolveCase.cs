using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using PluginXrm;

namespace UOP.Plugins.Case
{
    public class OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase : BasePlugin
    {
        private const string PostImageAlias = "Post";
        private const string Target = "Target";
        private const string NewEmail = "new_newemail";
        private const string ChangeEmailStatus = "new_changeemailstatus";
        private const string SubStatusReason = "new_substatusreason";
        private Guid _accountId;
       
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
                    // Check for update event
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase) && 
                        target.Attributes.Contains(ChangeEmailStatus) && target.Attributes[ChangeEmailStatus] != null)
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase));
                        localContext.ClearPluginTraceLog(nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase));

                        Entity postImageEntity = (localContext.PluginExecutionContext.PostEntityImages != null && 
                            localContext.PluginExecutionContext.PostEntityImages.Contains(PostImageAlias))
                            ? localContext.PluginExecutionContext.PostEntityImages[PostImageAlias]
                            : null;

                        int? newEmailStatus = postImageEntity?.GetAttributeValue<OptionSetValue>(ChangeEmailStatus).Value;
                        int confirmedOptionSetValue = new OptionSetValue(100000002).Value;

                        if (newEmailStatus != confirmedOptionSetValue)
                        {
                            return;
                        }

                        _accountId = postImageEntity.GetAttributeValue<EntityReference>("accountid").Id;
                        localContext.Trace($"Attempting to retrieve all active Email Change Request cases for account with id {_accountId}");

                        // Get all the active cases with title "Email Change Request" for the current account
                        using (var crmContext = new ServiceContext(localContext.OrganizationService) { MergeOption = MergeOption.NoTracking })
                        {
                            localContext.Trace($"Attempting to retrieve all active Email Change Request cases for account with id {_accountId}");
                            var allCases = (from incidents in crmContext.IncidentSet
                                            join account in crmContext.AccountSet on incidents.CustomerId.Id equals account.Id
                                            where incidents.CustomerId.Id == _accountId &&
                                                  incidents.SubjectId.Id.Equals(Guid.Parse(EmailChangeRequestSubjectId)) &&
                                                  incidents.StateCode == IncidentState.Active
                                            orderby incidents.CreatedOn descending
                                            select incidents)
                                            .ToList();

                            if (!allCases.Any())
                            {
                                localContext.Trace($"No active Email Change Request cases found for account with id {_accountId}");
                                return;
                            }

                            localContext.Trace($"Retrieved {allCases.Count} active Email Change Request cases for account id {_accountId}");

                            // Retrieve the latest active "Email Change Request" case for the current account and resolve it.
                            Incident latestCase = allCases.First();
                            ResolveCase(localContext, latestCase);

                            // Create new instance of account for update
                            var retrievedAccount = new Entity(Account.EntityLogicalName, _accountId);
                            var accountInstance = new Entity(Account.EntityLogicalName)
                            {
                                Id = retrievedAccount.Id,
                                Attributes =
                                {
                                    ["emailaddress1"] = latestCase.GetAttributeValue<string>(NewEmail),
                                    ["new_emailsetfromcase"] = true
                                }
                            };

                            localContext.Trace($"Updated account {accountInstance.GetAttributeValue<string>("name")}'s email to " +
                                               $"{latestCase.GetAttributeValue<string>(NewEmail)}");
                            localContext.OrganizationService.Update(accountInstance);

                            //Retrieve all other active "Email Change Request" cases (if any) for the current account that need to be closed as duplicate.
                            var casesToClose = allCases.Skip(1).ToList();
                            if (!casesToClose.Any())
                            {
                                localContext.Trace($"No other active cases for account with id {_accountId}");
                            }
                            else
                            {
                                // Close all other cases setting their built-in status and status reason to Cancelled
                                foreach (var currentCase in casesToClose)
                                {
                                    CloseCase(localContext, currentCase);
                                }

                                localContext.Trace($"Closed {casesToClose.Count} cases with status reason Cancelled for account with id {_accountId}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase)}. " +
                        $"Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        private void ResolveCase(ILocalPluginContext localContext, Incident latestCase)
        {
            // Change the case built-in status and status reason to Resolved and Problem solved respectively on the latest case created.
            // Change the Change Email Status to Approved on the latest case created. 
            latestCase[ChangeEmailStatus] = new OptionSetValue(100000004); //Approved
            latestCase[SubStatusReason] = new OptionSetValue(100000002); //Approved - Confirmed
            localContext.OrganizationService.Update(latestCase);

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

            localContext.OrganizationService.Execute(closeIncidentRequestResolved);
            localContext.Trace($"Set case status to Approved for Case with ID {latestCase.Id}");
        }

        private static void CloseCase(ILocalPluginContext localContext, Incident currentCase)
        {
            currentCase[ChangeEmailStatus] = new OptionSetValue(100000003); //Declined
            currentCase[SubStatusReason] = new OptionSetValue(100000001); //Cancelled - Duplicate
            localContext.OrganizationService.Update(currentCase);

            SetStateRequest setStateRequest = new SetStateRequest
            {
                EntityMoniker = new EntityReference(Incident.EntityLogicalName, currentCase.Id),
                State = new OptionSetValue((int)IncidentState.Cancelled), 
                Status = new OptionSetValue(6), 
            };

            localContext.OrganizationService.Execute(setStateRequest);
            localContext.TracingService.Trace($"Set case status to Cancelled for Case with ID {currentCase.Id}");
        }
    }
}
