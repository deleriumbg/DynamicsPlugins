# DynamicsPlugins

Dynamics Crm Plugins – Second Task

The Second Plugins Task would be as the following,
we need to create a plugin to be triggered on email address change event for Account entity.
On that plugin we need to skip the email change and keep the old email address value on the account record,
instead we need to create a case record connected to that Account (the case subject will be set to Email Change Request, new subject option is required for this purpose).
 
On the case record created we need to log the previous (pre value) and new email (post value) addresses (two new custom fields would be required here).
That case will have new option set field (named, Change Email Status) with the following option values: In Process, Confirmed, Declined and Approved, set to In Process by default.
Another plugin will be triggered on case creation with the following logic:
If the case is of email change type, we need to send an email message to the Account Person via his old email address, asking him to confirm that change.
- here there will be a missing logic (we need not implement it on this task), that email message should have a link to be clicked by the account person, once clicked certain logic will mark the case to be confirmed(we will do it manually this part).  
Third plugin will be registered on the case resolve event this time, looking for the value of new custom field named Change Email Status, if it is changed from In Process to Confirmed, we do the following:
Look for all Active Email change cases related to that account (account of the current source of trigger case), sort these cases in descending order, and get the last one created.
Close all other cases setting their built-in status and status reason to Cancelled.
change the Change Email Status to Approved on the latest case created.
Change that case built-in status and status reason to Resolved and Problem solved respectively on the latest case created.
Change the email address of the relevant account to the post email value on the latest case created
(the one just got approved).
Another fourth Plugin is required her on Account Update, check if the email address got changed (here you need to find a way if this change is based on a case or not).
If it was due to a case resolution, just send an email message to the Account telling him that his email address has been successfully changed and he can contact us through the new email address.
