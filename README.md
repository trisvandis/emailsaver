# emailsaver

About Email Saver:

Email saver is a macro for Outlook 2010 that automatically saves utility response and BS_Transmittals emails to the relevant work folders on the Underground Services shared network drive.
Email Saver works by extracting key information from the utility email or BS_transmittal and using it to find where to save the email on the shared network drive.  

How it works:

1. The macro first looks for the 6 digit Underground Services search reference in the subject title of the email.  It does this by searching the text for a 6 digit number beginning with either ‘LNE’ or ‘LNW’ or ‘SET’ or ‘WES’ or ‘SCT’.  In the case of saving individual emails (i.e. one email selected) it can’t find this it will prompt the user to manually enter the search reference via a user entry form.
2. When the folder reference has been identified, the macro will loop through each of the active user references (e.g. ‘TB5’) listed in an array.  It uses each user reference to construct a folder path (e.g. ‘O:\Buried Services\BSRM\TB5\LNE100000’) and tests if the folder path exists.  If the folder path doesn’t exist, it will move on to the next user reference in the array, construct a path and test it.  It will keep doing this until it either finds a folder path that does exist (in which case it will move onto the next step) or if it fails to identify a user that matches the search reference, in the case of individual selections,  it will prompt the user to manually enter the user reference.
3. The macro checks the email senders email address in order to work out what utility code to save the email and attachments under.  In the case of saving individual emails, the macro can’t determine what utility code it is, it will prompt the user to manually enter the utility code.
4. Once the macro has identified the folder path (my matching the correct user ref with the search ref) and has obtained the utility code to save it under, it will then print a pdf copy of the email into the folder path (NB:  there are certain rules set up for utility codes such as T02 where the email does not need to be saved – in these cases this step is skipped).
5. The macro will then process any relevant attachments and save them as PDFs to the correct folder path.  Email Saver has been set up so it doesn’t save specific attachments, such as the Vodafone safety file, as these are automatically copied into the user folder by the USWMS database once the file has been saved and the ‘check PDFs’ query has been run.  It also automatically converts word documents to pdf.
6. The macro will finally display a report on what has been saved.  In the case of multiple emails, it will flag up any that it has been unable to process.


First Time Setup:

In order to setup and start using Email Saver, follow these steps:

1. Open Outlook and press Alt + F11.  This will open the ‘Microsoft Visual Basic for Applications’ window as pictured below:
 
2. Go to the following folder on the O Drive:  O:\Buried Services\BSRM\!Request Forms etc\Weekly Workload\Outlook Macros\Email Saver and open the folder with the most recent version.
Drag all of the macro’s modules (i.e. all files within the folder) into the ‘Project’ area, located in the top left portion of the work area.
 
3. Select Tools > References from the menu bar. In the “References” pop-up menu, tick the box next to’ Microsoft Word 14.0 Object Library’.

4. Close the ‘Microsoft Visual Basic for Applications’ window (remember to click ‘yes’ to save the changes) and return to your Outlook inbox.  

5. Once back in Outlook, click on the ‘File’ Tab and select ‘Options’.
 
6. In the ‘Outlook Options’ pop-up menu, select ‘Customize Ribbon’
 
7. In the Main Tab section to the right of the ‘’Outlook Options’ menu, click on Home(Mail) so that it is highlighted:
 
8. With the Home(Mail) main tab highlighted, click the ‘New Group’ button near bottom of the pop up.

9. A ‘New Group (Custom)’ will be added within your Home(Mail) main tab. Select ‘New Group (Custom)’ and click ‘Rename’.
 
10. Rename the group ‘Saving’ and click ‘Ok’ to close the pop-up.
 
11.  As pictured, reposition the ‘Saving (Custom)’ group in the list, so it falls between the  ‘New’ and ‘Delete’ groups.
 
You can use these buttons to move the group up and down the list: 
 
12. In the ‘Choose Command from’ drop-down menu, select ‘Macros’ from the list of options.

13. The macro will be displayed in the window below the drop down.  Select the option named ‘Email_Saver.StartEmailSaver’ macro and add it to the Saving group you created: 

14. Once the macro has been added to the ‘Saving’ group on the right-hand side of the screen, ensure you have the macro that you’ve just moved across highlighted and click the ‘Rename’ button. 

15. In the ‘Rename’ pop-up window, enter ‘Email Saver’ as the ‘Display Name’, and change the icon to the one highlighted below. Once this has been done, click ‘Ok’ to close the pop-up:
 
16. The ‘Saving(Custom)’ group should now look like this:
 
17. Click ‘Ok’ to close the ‘Outlook Options’ pop-up window. Return to your ‘Home’ tab in Outlook.  You should now be able to see the newly created ‘Email Saver’ button in the ribbon. 
 
18. Email Saver is now setup and ready to use!

Saving Individual Emails:
1. To use Email Saver on individual emails, (a.) select the uemail you want to save in outlook and (b.) click the Email Saver button: 

2. The Email Saver script will run and save the email and any relevant attachments. Once, complete a pop up box will appear confirming the file has been saved. Click ‘Ok’ to dismiss this.

3. The email will automatically be marked as read and a tick will appear in the ‘Flag Status’ box.
 
Notes on individual saving:
•	The macro automatically works out the folder reference and utility code to save it under using information contained within the email.
•	If the macro can’t find specific information such as the folder reference, user reference or utility type (e.g. T02) it a form will appear asking the user to manually enter the missing information.
•	There are three types of manual entry forms that may appear if information is missing.  These are as follows:

Enter Search Ref (appears if the macro can’t find the search reference in the email subject) – Use the radio buttons to select the territory and enter the 6-digit reference from USMWS into the text box. Click ‘OK’ once you have entered all the information to proceed to the next step or ‘cancel’ to abort saving the email.
 
Enter User Ref (appears if the macro can’t find a user ref that corresponds to the search reference):
Select the user that corresponds to where the folder is saved.  Alternatively you can use the ‘Enter Manually’ text box to enter alternative users that are not listed on the form itself. Click ‘OK’ to continue or ‘Cancel’ to abort saving the email.

 
Enter Utility code (appears if the macro can’t determine the utility code from the senders email address)  - Select the utility code that the email corresponds to.  Alternatively you can manually enter the utility code in the ‘Enter Manually’ text box.  Click ‘OK’ to continue or ‘Cancel’ to abort saving the email.
 
	
Saving Multiple Utility Emails:
1. You can also email saver to save multiple emails in one go.  In order to do this, hold down the ‘Shift’ key and highlight all the emails you want to save.
 
2. Click the Email Saver button.  The macro will run through each email highlighted and save them along with any attachments to the relevant folders. When the macro has finished process the highlighted email it will display a Confirmation Report.  Click ‘OK’ to dismiss this.
 
Notes on Multiple saving:
 If the marco can’t save any of the selected emails it will mark them as ‘Failed to save’ on the Confirmation Report.  
 
It will also assign a red flag and a red category to them in outlook…
 
Updating Email Saver – All Users

If a new version of Email Saver is released, follow these steps to update your version:

1. Open Outlook 2010 and press Alt + F11.  This will open the ‘Microsoft Visual Basic for Applications’ window as pictured below:
 
2. Select each of the files listed in the ‘modules’ folder and right-click the ‘Remove …” option. When prompted if you want to export before removing, click ‘No’.
 
3. Go to the following folder on the O Drive:  O:\Buried Services\BSRM\!Request Forms etc\Weekly Workload\Outlook Macros\Email Saver and open the folder with the most recent version.
Drag all of the macro’s modules (i.e. all files within the folder) into the ‘Project’ area, located in the top left portion of the work area.
4.  Close the ‘Microsoft Visual Basic for Applications’ window (remember to click ‘yes’ to save the changes) and return to your Outlook inbox.  Email Saver should now be updated.  Happy Saving! 


