<# Your personnal token from https://developer.webex.com/docs/bots #>
$token = "REPLACE_WITH_YOUR_PERSONNAL_TOKEN"

<# Create folder if down't exist #>
$FolderName = "./fetched-data\"
if (Test-Path $FolderName) {   
    Write-Host "Folder exists" -ForegroundColor Cyan
    Write-Host "Old files were deleted" -ForegroundColor Cyan
    Get-ChildItem –Path  ./fetched-data/ –Recurse -include *.txt | Where-Object { $_.CreationTime –lt (Get-Date).AddMinutes(-5) } | Remove-Item
}
else {  
    New-Item $FolderName -ItemType Directory
    Write-Host "Folder created successfully" -ForegroundColor Cyan
}

<# Generate random ID for temp file naming #>
$ID = Get-Random -Maximum 100

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

<# Starting script #>
Write-Host "Welcome to Webex AvailaBOT !" -ForegroundColor Green
$email = Read-Host -Prompt "Enter the user's e-mail you'd like to get a notification when available"
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host "The entered e-mail is" $email -ForegroundColor Green

<# API call parameters #>
$Header = @{
    "authorization" = "Bearer $token"
}
$Parameters = @{
    Method      = "GET"
    Uri         = "https://webexapis.com/v1/people?email=" + $email
    Headers     = $Header
    ContentType = "application/json; charset=utf-8"
    OutFile     = "./fetched-data/" + $ID + ".txt"
}

<# Fetch once Cisco's API to get user's data and current status #>
Invoke-RestMethod @Parameters
$data = (Get-Content ./fetched-data/$ID.txt -Encoding UTF8) | ConvertFrom-Json
$UserStatus = $data.items.status
$UserFirstName = $data.items.firstName
$UserLastName = $data.items.lastName
Write-Host " "
Write-Host("Fetching $UserFirstName $UserLastName's status") -ForegroundColor Green

<# Fetch Cisco's API each 30 seconds #>
Do {
    <# Fetch user's data from Cisco API #>
    Invoke-RestMethod @Parameters
    $data = (Get-Content ./fetched-data/$ID.txt -Encoding UTF8) | ConvertFrom-Json
    $UserStatus = $data.items.status

    if ($UserStatus -ne "active") {
        <# Display the user's current status #>
        Write-Host($UserFirstName + " " + $UserLastName + " is " + $UserStatus) -ForegroundColor Green

        <# Create a progress bar to whow the remaining time before next check #>
        $seconds = 30
        1..$seconds |
        ForEach-Object { $percent = $_ * 100 / $seconds; 
            Write-Progress -Activity "Fetching" -Status "Next check in $($seconds - $_) seconds..." -PercentComplete $percent;
            Start-Sleep -Seconds 1
        }    
    }

} While ($UserStatus -ne "active")

<# Hide progress bar once the user is available #>
Write-Progress -Completed -Activity "Fetching"

<# Preparing the toast notification and closing the script window #>
$NotificationText = ($UserFirstName + " " + $UserLastName + " est désormais disponible dans Webex !")
Write-Host ($UserFirstName + " " + $UserLastName + " is currently " + $UserStatus) -ForegroundColor Green
Show-Notification -ToastTitle "Webex availaBOT" -ToastText $NotificationText

<# Create a progress bar to whow the remaining time before closing windows #>
Write-Host "This window will close automatically in 10 seconds..." -ForegroundColor Cyan
$seconds = 10
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
    Write-Progress -Activity "Closing" -Status "Closing window in $($seconds - $_) seconds..." -PercentComplete $percent;
    Start-Sleep -Seconds 1
}

<# Cleaning old files #>
Write-Host "Temporary files are being cleaned..." -ForegroundColor Cyan
Remove-Item ./fetched-data/$ID.txt 
Get-ChildItem –Path  ./fetched-data/ –Recurse -include *.txt | Where-Object { $_.CreationTime –lt (Get-Date).AddMinutes(-5) } | Remove-Item