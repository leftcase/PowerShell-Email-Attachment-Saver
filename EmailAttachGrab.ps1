$filepath = "H:\test\" # <--- change me! Your photos end up here! Path must exist.

# Invoke the Outlook API to get access to the messaging namespace
# This gives us access to Outlook 'stuff' as objects. For reference see below:
# https://msdn.microsoft.com/VBA/Outlook-VBA/articles/object-model-outlook-vba-reference
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

# Where's the inbox?
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

#$inbox.Items returns an items collection object from the inbox
foreach ($item in $inbox.Items)
    {
    if ($item.UnRead -eq 'True') # if it ain't read yet - proceed
        {$item.Subject # testing-remove
        # Parse the item's attachments (if there are none, this does nowt)
        $item.Attachments | foreach ($_.Filename) {$_.SaveAsFile((Join-Path $filepath ($item.ConversationID + '-' + $_.Index + '-' + $_.FileName)))}
        $item.Unread = $false # Mark the mail item as read
         }
    } 
