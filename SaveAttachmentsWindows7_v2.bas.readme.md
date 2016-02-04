ABOUT THE MACROS

Email messages with attached files can take up a lot of space in your mailbox. These Outlook macros can help with this problem. The SaveAttachments macro will save your message attachments to a folder you specify and will insert a note in the email telling you what files were saved. Each saved file will have a hyperlink that can be clicked on to open the saved file. For example:

[Attachments saved: 
    D:\DCData\Projects\DC Preservation Network\PWG\HUD Propose PBV HERA Rules.pdf ] 

The DeleteAttachments macro will delete all attachments in a message and insert a similar note listing what files were deleted. 

[Attachments deleted: 
    HUD Propose PBV HERA Rules.pdf ] 

I use SaveAttachments primarily for emails that are sent TO ME, since I will want to save those attachments to my hard disk. I use DeleteAttachments for emails that I send TO OTHERS (i.e., messages in the Sent Items folder), since I presumably already have a copy of those attachments on my hard disk.

While these macros are "use at your own risk," I've been using them for several years now and they have not given me any problems. 

The macros will preserve the formatting of HTML messages. Plain text or rich text messages may have their formatting altered by the macros. 

The macros have to be installed before they can be used for the first time. Once they are installed, they will available to you in all future Outlook sessions. But, as described below, you may have to enable macros for each time you restart Outlook.

UPDATING THE MACROS (IF OLDER VERSIONS ALREADY INSTALLED)

1.	In Outlook, click on the Developer tab and then click the Visual Basic button. This should open up the Visual Basic Editor window. 
2.	In the Project window on the left hand side, double-click on the Modules folder (under Project1) so that the Module2 object appears in the tree. 
3.	Right-click on Module2 and select “Remove Module2…” from the popup menu. Click No when asked if you want to export the module first. 
4.	Go to File > Import File and navigate to K:\Metro\PTatian\CENTER\ in the dialog. 
5.	Choose the file SaveAttachmentsWindows7_v2.bas and click Open. 
6.	Save the "project" by pressing Ctrl-S or selecting File > Save from the menu. 
7.	Finally, close the Visual Basic Editor (File > Close and Return to Microsoft Outlook). 

You may also need to reset your security settings to allow you to run the macros. 

1.	In Outlook, click on the Developer tab and then click the Macro Security button. 
2.	In the dialog box, under “Macro Settings,” select “Notifications for all macros,” and click OK. 
3.	You may need to close Outlook and restart for the new settings to take effect.

INSTALLING THE MACROS (FOR THE FIRST TIME)

First, you need to display the Developer tab in your Outlook ribbon. 

1.	Click on the File tab in the Outlook window and select Options. 
2.	In the "Outlook Options" box, click on "Customize Ribbon," and check the box next to Developer in the "Customize the Ribbon" list on the right. 
3.	Click OK. The Developer tab should now be visible above the ribbon.

Next, you need to enable macros in Outlook. 

1.	Click on the Developer tab above the ribbon and then click on "Macro Security." 
2.	Under "Marco Settings," select "Notifications for all macros." Click on OK. 
3.	You will have to exit Outlook completely and then restart it for this change to take effect.

Finally, install the macros.

1.	Back in Outlook, click on the Developer tab and then click the Visual Basic button. This should open up the Visual Basic Editor window. 
2.	Go to File > Import File and navigate to K:\Metro\PTatian\CENTER\ in the dialog. 
3.	Choose the file SaveAttachmentsWindows7_v2.bas and click Open. 
4.	Save the "project" by pressing Ctrl-S or selecting File > Save from the menu. 
5.	Close the Visual Basic Editor (File > Close and Return to Microsoft Outlook). 

Once installed, the macros will available in all future Outlook sessions. 

USING THE MACROS

In Outlook, open a mail message with attachments. Run the SaveAttachments or DeleteAttachments macro either from the message window ribbon command (Developer > Macros) or by pressing Alt-F8. Highlight the macro you wish to use in the dialog and click Run.

Note:  When you first run a macro in an Outlook session, Outlook may ask you if you want to enable macros. Click the Enable button. Macros will remain enabled until you close and next reopen Outlook. You must re-enable macros each time you restart Outlook. If you are not asked if you want to enable macros, you may need to adjust your security settings. In Outlook, click on the Developer tab and then click the Macro Security button. In the dialog box, under “Macro Settings,” select “Notifications for all macros,” and click OK. You may need to close Outlook and restart for the new settings to take effect.

For the SaveAttachments macro, a dialog box should appear prompting you for the folder where you want the attachments saved. Unfortunately, I could not figure out how to make an "Open File" dialog appear so you have to either type in the folder path or go to Windows Explorer and copy and paste the full folder path from the Address bar. Click OK. (Each subsequent time you run the macro in the same Outlook session, the macro will remember the folder where you last saved attachments.)

Each of the attachments in the message will be saved to the folder you specified. If a copy of the file already exists in the folder, you will be asked if you want to replace it. If you say “No,” that attachment will be skipped. Any attachments that are saved will automatically be deleted from the email message. A list of saved attachments will be inserted at the end of the message. 

Once the macro is finished, the message will be in an "unsaved" state. If you try to close the message window or move to the next or previous message with the arrow buttons, you will be asked if you want to save the message first. You should say “Yes,” unless you want to undo the macro's work. Saying “No” will return the message to its original state, with all attachments intact. (But they will still be saved on your hard disk.)  Note, however, that if you attempt to move the message to another folder, Outlook will save the message automatically without asking you. (Helpful Outlook!)

The DeleteAttachments macro works very similarly. When you run it, a prompt will appear to confirm that you want to delete all attachments in the open email. If you click “Yes,” all attachments will be deleted and a list of deleted attachments will be inserted at the end of the message. (These attachments are NOT saved to your hard disk!) Again, once you run the macro the message will be in an "unsaved" state. At this point, you can still recover the deleted attachments by closing the message and saying "No" when asked "Do you want to save changes?"  Saying "Yes" will save the message with attachments deleted - and the attachments will be gone! The message with deleted attachments will also be saved automatically if you move the message to another folder.

Note:  These macros only work on emails that have been saved. That is, if you make a change to the email in some way while it is open and run the macro, the macro will refuse to do its thing until you resave the message. You can either hit the Save button on the open message, or close the message, saying “No” when asked to save changes, and reopen it before running one of the macros.

UPDATED 5/6/14

