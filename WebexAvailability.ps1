<# Your personnal token from https://developer.webex.com/docs/bots#>
$token = "REPLACE_WITH_YOUR_PERSONNAL_TOKEN"

<# Toast notification function #>
function Show-Notification {
    [cmdletbinding()]
    Param (
        [string]
        $ToastTitle,
        [string]
        [parameter(ValueFromPipeline)]
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text | where { $_.id -eq "1" }).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
    ($RawXml.toast.visual.binding.text | where { $_.id -eq "2" }).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "PowerShell"
    $Toast.Group = "PowerShell"
    $Toast.Priority = 1
    $Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(500)
    $Toast.SuppressPopup = $false
    $Toast.ExpiresOnReboot = $false

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Webex availaBOT")
    $Notifier.Show($Toast);
}

<# Greetings text and requesting user input #>
Write-Host "Welcome to Webex AvailaBOT !" -ForegroundColor Green
$email = Read-Host -Prompt "Enter the user's e-mail you'd like to get a notification when available"
Write-Host "The entered e-mail is" $email -ForegroundColor Green

<# API call parameters #>
$Header = @{
    "authorization" = "Bearer $token"
}

$Parameters = @{
    Method      = "GET"
    Uri         = "https://api.ciscospark.com/v1/people?email=" + $email
    Headers     = $Header
    ContentType = "application/json"
}

<# Fetching user's current status, first and last name #>
$APICall = Invoke-RestMethod @Parameters
$UserStatus = $APICall.items.status
$UserFirstName = $APICall.items.firstName
$UserLastName = $APICall.items.lastName

<# Notifying current status #>
Write-Host ("Fetching " + $UserFirstName + " " + $UserLastName + "'s status") -ForegroundColor Green

<# If the status is not available, fetch and check it every 10 seconds #>
Do {
    $APICall = Invoke-RestMethod @Parameters
    $UserStatus = $APICall.items.status

    if ($UserStatus -ne "active") {
        Write-Host "The user is currently" $APICall.items.status "- Retrying in 10 seconds..." -ForegroundColor Green
        Start-Sleep -Seconds 10
    }

} While ($UserStatus -ne "active")

<# Notify that the requested user is online #>
Write-Host "!!! The user is currently" $APICall.items.status -ForegroundColor Green

<# Custom text for the Toast notification #>
$NotificationText = ($UserFirstName + " " + $UserLastName + " est désormais disponible dans Webex !")
Show-Notification -ToastTitle "Webex availaBOT" -ToastText $NotificationText

<# Automatically close the script windows and notify the user abtout it #>
Write-Host "This window will close automatically in 10 seconds..." -ForegroundColor Green
Start-Sleep -Seconds 10