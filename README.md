# WebexTeams-AvailaBOT

- [WebexTeams-AvailaBOT](#webexteams-availabot)
  - [AvailaBOT in a nutshell](#availabot-in-a-nutshell)
  - [How to use it](#how-to-use-it)
  - [How to set it up](#how-to-set-it-up)
  - [Launching the script](#launching-the-script)

## AvailaBOT in a nutshell

This PowerShell script will notify you when a contact is available in Webex Teams, for a chat of a call.

## How to use it

Simply double click on  `Run availaBOT.bat` and follow the instructions:

![image](images/1.%20Homescreen.png?raw=true "Script launched")

Simply enter the e-mail adress of the colleague or user you want to be notified and press enter. The script will check the user's status every ten second.
When the user is available, you'll get a Windows notification.

## How to set it up

In order to get the script to work, you'll need to create a token to call Cisco's API. For that, simply go to [this page](https://developer.webex.com/docs/bots) and click `Create a Bot`.
After filling some details, you'll get a Bearer token.

Simply paste this token on line 2 in `WebexAvailability.ps1`:

```powershell
<# Your personnal token from https://developer.webex.com/docs/bots #>
$token = "REPLACE_WITH_YOUR_PERSONNAL_TOKEN"
```

and you're good to go !

## Launching the script

You can start the script by:

1. Right-clicking on `WebexAvailability.ps1` and selecting `Execute with PowerShell`
2. Double clicking on `Run availaBOT.bat` (preferred)

For the second option, you can even create a shortcut to this file and place it anywhere you want.
For example, I placed mine in the Start Menu, that way, I can quickly start the script anywhere:

![image](images/Start%20menu.png?raw=true "Start menu")
