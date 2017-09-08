# PowerShell-Email-Attachment-Saver
Use PowerShell to copy email attachments to a folder

Open Outlook first. This script accesses the logged on users’ mailbox. Don’t run it with administrative privileges...

I’ve tested this against my mailbox with a couple of thousand emails. It seems to copy the attachments and didn’t set anything on fire. Original emails are kept in place and marked as read when processed. Bear in mind though that I’m running Windows 10 and Outlook 2013 so I wouldn’t read anything into that lol. Read the comments to understand what the script is doing…

Each attachment is copied to a network drive of your choice (see below). A unique filename is created to ensure that the attachment name is globally unique and an incremental number is added to handle emails with multiple file attachments with identical names (yes, people can do this). The filename is formatted as such {unique email conversation ID}-{attachment index number starting from 1}-{original attachment filename} 
