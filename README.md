# DynamicsPlugins

Dynamics CRM Plugins Tasks

Part 1:

Create a plugin to be triggered on email address change event for Account entity.
On that plugin we need to skip the email change and keep the old email address value on the Account record,
instead we need to create a Case record connected to that Account (the case subject will be set to Email Change Request,
new Subject option is required for the purpose).
On the Case record created we need to log the previous (pre value) and new Email (post value) addresses (two new custom fields would be required here).
That Case will have new option set field (named Change Email Status) with the following option values: 
In Process, Confirmed, Declined and Approved, set to In Process by default.

Another plugin will be triggered on Case creation with the following logic:
If the Case is of Email Change Request type, we need to send an Email message to the Account Person via his old Email address, 
asking him to confirm that change.
Here there will be missing logic (we don't need to implement it on this task), that Email message should have a link to be clicked by the account person, once clicked certain logic will mark the case to be confirmed (we will do it manually).  

Third plugin will be registered on the Case resolve event, this time looking for the value of new custom field named Change Email Status. If it is changed from In Process to Confirmed, we do the following:
Look for all Active Email Change Request Cases related to that Account (Account of the current source of trigger Case), 
sort these Cases in descending order, and get the last one created.
Close all other Cases setting their built-in Status and Status Reason to Cancelled.
Change the Change Email Status to Approved on the latest Case created.
Change that Case built-in Status and Status Reason to Resolved and Problem solved respectively on the latest Case created.
Change the Email address of the relevant Account to the post Email value on the latest Case created
(the one just got approved).

Another fourth Plugin is required on Account Update, check if the Email address got changed 
(here you need to find a way if this change is based on a Case or not).
If it was due to a Case resolution, just send an Email message to the Account telling him that his Email address has been successfully changed.

Part 2:

Re-implement the same functionalities using two custom workflows.
One for Account Email address initial update, creating a change Email request Case.
The second one should be triggered on Case Status change to Confirmed and execute the rest of the functionality.
