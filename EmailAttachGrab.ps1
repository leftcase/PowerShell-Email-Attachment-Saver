<# 
Copyright (C) 2017 Chris Rowson - christopherrowson@gmail.com
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
#>

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
