using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;

namespace UOP.Plugins.Case
{
    public class OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase : BasePlugin
    {
        private const string PreImageAlias = "Pre";
        private const string PostImageAlias = "Post";
        private const string Target = "Target";
        private Guid _accountId;

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
                    //If it's update get the latest data 
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase) && 
                        target.Attributes.Contains("new_changeemailstatus") && target.Attributes["new_changeemailstatus"] != null)
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase));
                        localContext.ClearPluginTraceLog(localContext, nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase));

                        //Obtain the pre image and post image entities
                        Entity preImageEntity = (localContext.PluginExecutionContext.PreEntityImages != null && localContext.PluginExecutionContext.PreEntityImages.Contains(PreImageAlias)) ?
                            localContext.PluginExecutionContext.PreEntityImages[PreImageAlias] : null;
                        Entity postImageEntity = (localContext.PluginExecutionContext.PostEntityImages != null && localContext.PluginExecutionContext.PostEntityImages.Contains(PostImageAlias)) ?
                            localContext.PluginExecutionContext.PostEntityImages[PostImageAlias] : null;

                        int? newEmailStatus = postImageEntity?.GetAttributeValue<OptionSetValue>("new_changeemailstatus").Value;
                        int confirmedOptionSetValue = new OptionSetValue(100000002).Value;

                        if (newEmailStatus != confirmedOptionSetValue)
                        {
                            return;
                        }

                        //Get all the active cases with title "Email Change Request" for the current account
                        ServiceContext crmContext = new ServiceContext(localContext.OrganizationService)
                        {
                            MergeOption = MergeOption.NoTracking
                        };

                        _accountId = postImageEntity.GetAttributeValue<EntityReference>("accountid").Id;
                        localContext.Trace($"Attempting to retrieve all active Email Change Request cases for account with id {_accountId}");

                        List<Incident> allCases = (from incidents in crmContext.IncidentSet
                                                   join account in crmContext.AccountSet on incidents.CustomerId.Id equals postImageEntity.GetAttributeValue<EntityReference>("accountid").Id
                                                   where incidents.CustomerId.Id == postImageEntity.GetAttributeValue<EntityReference>("accountid").Id &&
                                                         incidents.Title.Equals("Email Change Request") &&
                                                         incidents.StateCode == IncidentState.Active
                                                   orderby incidents.CreatedOn descending
                                                   select incidents)
                                        .ToList();

                        if (!allCases.Any())
                        {
                            localContext.Trace($"No active Email Change Request cases found for account with id {_accountId}");
                            return;
                        }

                        localContext.Trace($"Retrieved {allCases.Count()} active Email Change Request cases for account id {_accountId}");

                        Incident latestCase = allCases.First();
                        List<Incident> casesToClose = allCases.Skip(1).ToList();

                        if (!casesToClose.Any())
                        {
                            localContext.Trace($"No other active cases for account with id {_accountId}");
                        }
                        else
                        {
                            //Close all other cases setting their built-in status and status reason to Cancelled
                            foreach (var currentCase in casesToClose)
                            {
                                currentCase["new_changeemailstatus"] = new OptionSetValue(100000003);
                                localContext.OrganizationService.Update(currentCase);

                                SetStateRequest setStateRequest = new SetStateRequest
                                {
                                    EntityMoniker = new EntityReference("incident", currentCase.Id),
                                    State = new OptionSetValue(2),
                                    Status = new OptionSetValue(6),
                                };

                                localContext.OrganizationService.Execute(setStateRequest);
                                localContext.TracingService.Trace($"Set case status to Cancelled for Case with ID {currentCase.Id}");
                            }

                            localContext.Trace($"Closed {casesToClose.Count()} cases with status reason \"Cancelled\" for account with id {_accountId}");
                        }

                        //Change the case built-in status and status reason to Resolved and Problem solved respectively on the latest case created.
                        //Create Incident Resolution
                        var incidentResolutionResolved = new IncidentResolution
                        {
                            Subject = "Case Resolved",
                            IncidentId = new EntityReference(Incident.EntityLogicalName, latestCase.Id),
                            ActualEnd = DateTime.Now
                        };

                        //Close Incident
                        var closeIncidentRequestResolved = new CloseIncidentRequest
                        {
                            IncidentResolution = incidentResolutionResolved,
                            Status = new OptionSetValue(5)
                        };

                        localContext.OrganizationService.Execute(closeIncidentRequestResolved);

                        //Change the Change Email Status to Approved on the latest case created. 
                        latestCase["new_changeemailstatus"] = new OptionSetValue(100000004);
                        localContext.Trace($"Set case status to Approved for Case with ID {latestCase.Id}");

                        Entity accountObj = localContext.OrganizationService.Retrieve("account", _accountId, new ColumnSet(true));
                        accountObj.Attributes["emailaddress1"] = latestCase.GetAttributeValue<string>("new_newemail");
                        accountObj.Attributes["new_emailsetfromcase"] = true;
                        localContext.Trace($"Updated account {accountObj.GetAttributeValue<string>("name")}'s email to {latestCase.GetAttributeValue<string>("new_newemail")}");
                        localContext.OrganizationService.Update(accountObj);
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnCaseChangeEmailStatusConfirmed_OnPostUpdateEmailAndResolveCase)}. Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
