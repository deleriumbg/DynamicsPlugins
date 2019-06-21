using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using PluginXrm;

namespace UOP.Plugins.Account
{
    public class OnUpdateAccountEmailFromCaseResolution_SendEmailToAccountNewMail : BasePlugin
    {
        private const string Target = "Target";
        private Guid _userId;
        private Guid _emailId;
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
                    //Check for update event for account with email set from case 
                    if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase) && target.GetAttributeValue<bool>("new_emailsetfromcase"))
                    {
                        localContext.Trace("Entered {0}.Execute()", nameof(OnUpdateAccountEmailFromCaseResolution_SendEmailToAccountNewMail));
                        localContext.ClearPluginTraceLog(localContext, nameof(OnUpdateAccountEmailFromCaseResolution_SendEmailToAccountNewMail));
                        localContext.Trace($"Attempting to send email to user informing him about his email change...");

                        // Get a system user to send the email (From: field)
                        WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                        WhoAmIResponse systemUserResponse = (WhoAmIResponse)localContext.OrganizationService.Execute(systemUserRequest);
                        _userId = systemUserResponse.UserId;

                        //Get the email recipient user(Account entity)
                        Entity account = localContext.OrganizationService.Retrieve("account", target.GetAttributeValue<Guid>("accountid"), new ColumnSet(true));
                        if (account == null)
                        {
                            localContext.Trace($"Unable to retrieve the account. Aborting plugin execution.");
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
                            PartyId = new EntityReference("account", _accountId),
                            AddressUsed = target.GetAttributeValue<string>("emailaddress1")
                        };

                        localContext.Trace("Created To and From activity parties.");

                        // Create an email message entity
                        Email email = new Email
                        {
                            To = new ActivityParty[] { toParty },
                            From = new ActivityParty[] { fromParty },
                            Subject = "Email successfully changed",
                            Description = $"Hello, {account.GetAttributeValue<string>("name")}.{Environment.NewLine}" + 
                                          $"We are writing you to inform you that your primary email was successfully changed " +
                                          $"to {account.GetAttributeValue<string>("emailaddress1")}.{Environment.NewLine}" + 
                                          $"Thank you for using our services!",
                            DirectionCode = true
                        };
                        _emailId = localContext.OrganizationService.Create(email);
                        localContext.Trace($"Created {email.Subject} email with description: {email.Description}.");

                        SendEmailRequest sendEmailRequest = new SendEmailRequest
                        {
                            EmailId = _emailId,
                            TrackingToken = "",
                            IssueSend = true
                        };

                        account.Attributes["new_emailsetfromcase"] = false;
                        localContext.OrganizationService.Update(account);
                        localContext.Trace($"Reset the Email Set From Case field to default value");

                        SendEmailResponse sendEmailResponse = (SendEmailResponse)localContext.OrganizationService.Execute(sendEmailRequest);
                        localContext.Trace($"Sent the email successfully changed message to user {account.GetAttributeValue<string>("name")} with id {_accountId}.");
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"An error occurred in {nameof(OnUpdateAccountEmailFromCaseResolution_SendEmailToAccountNewMail)}. Exception details: {ex.Message}");
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
