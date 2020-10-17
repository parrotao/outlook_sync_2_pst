# outlook_sync_2_pst


Hi all:

   This tools is for download huge email from outlook to local PST file through outlook VBS.
   
   The orginal request is from my company to try to reduce O365 mail box size due to migrate to a new o365 org.

   I am not a programer, it is just a tool to make activity effective.
   
   thanks for your interesting.
   
Regards

parrotao@gmail.com


# How to install

1. Open outlook
2. Press Alt-F11
3. Right Click on "Project" tab 
4. import UserForm1.frm & Module1.bas
5. Double click on UserForm1.frm & Press F5

# How to use
1. Click "load Mailbox from Outlook"
2. Choose the mail box  which you want to download from
3. Click "Set From"
4. Choose the detail folder which need to download from
5. ReChoose the mail box which you want to download to
6. Click "Set To"
7. Choose the detail folder which need to download to 
8. Select From time and To time through you input now-[%day%] 
9. Click go

Advance function
it is not recommnaded for no-IT users!!!

it is for huge folders which keep email several years.

1. Create PST file with name Mailbox_YYYY for example Mailbox_2020, Mailbox_2019
2. Create **same folder structure** in **each** PST file
3. Run this form
4. Click "load Mailbox from Outlook"
5. Choose the mail box  which you want to download from
6. Click "Set From"
7. Choose the detail folder which need to download from
8. Click "ONLY for Helpdesk"
* it will download email to different PST based on differnet recieved year and keep this year email in the orginal mailbox.

# Troubleshooting
1. ensure VBS Addin is enabled
2. enable Marco


