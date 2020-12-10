# RDS-Mgmt_using_ConnectionBroker.ps1
Manage Windows 2012 R2 Remote Desktop sessions using a GUI &amp; Connection Broker


With Windows Server 2012 R2 your companies helpdesk does not have any possibility to manage Remote Desktop User Sessions as there is simply no tool to accomplish that on Windows 2012 R2.

This is where this script comes into play as a replacement for the Remote Desktop Services Manager of Windows 2008.

What it does:

 
Connect to the Connection broker and import all Session Collections
After choosing a collection it imports all user sessions in this collection
Letting you choose one or more user sessions, you can:
send a message to these users
logoff the user from a session
disconnect a user from the session
mirror the user session
 

This script is based on the script of Denis Cooper (thank you very much for your work!).

There are some prereqs for this to work:

Your helpdesk (very unfortunately) needs to have admin rights on the connection broker (i didn't find a way to delegate a kind of "query connection broker right").
you need to set the following on each of your Remote Desktop Session Host(s) to the helpdesk group:
wmic /namespace:\\root\CIMV2\TerminalServices PATH Win32_TSPermissionsSetting WHERE (TerminalName ="RDP-Tcp") CALL AddAccount "<domain\group>",2

Note: Figured out that you have to REBOOT your RDSH(s) after this step!

Replace "<Your_RDSH_goes_here>" at the end of the script with FQDN(!) of your connection broker
in order to mirror a session you need an up-to-date Remote Desktop Client
 

Hope you enjoy it!

I published another version of this script where helpdesk needs no admins rights on the connection broker server. Of course, that version does not support multiple collections as it is host based.

 

Disclaimer:

This posting is provided "AS IS" with no warranties or guarantees!

I assume no liability for damage by the supplied code!
