<# Your personnal token from https://developer.webex.com/docs/bots #>
$token = "REPLACE_WITH_YOUR_PERSONNAL_TOKEN"

<# Clear the current windows #>
Clear-Host

<# Gloabl var #>
$global:email = $null
$global:heroIMG = Get-ChildItem ./icons/hero.png
$global:openicon = Get-ChildItem ./icons/open.png
$global:closeicon = Get-ChildItem ./icons/close.png
$global:avatar = $null

<# Create folder if down't exist #>
$FolderName = "./fetched-data\"
if (Test-Path $FolderName) {   
    Write-Host "Folder exists" -ForegroundColor Cyan
    Get-ChildItem –Path ./fetched-data/ –Recurse -include *.txt, *.png | Where-Object { $_.CreationTime -lt (Get-Date).AddMinutes(-5) } | Remove-Item
    Write-Host "Old files were deleted" -ForegroundColor Cyan
}
else {  
    New-Item $FolderName -ItemType Directory
    Write-Host "Folder created successfully" -ForegroundColor Cyan
}

<# Generate random ID for temp file naming #>
$ID = Get-Random

<# Toast notification function #>
function Show-Notification($NotificationText) { 

    # Required
    $winTitle = "Webex AvailaBOT"
    $audSource = "ms-winsoundevent:Notification.Looping.Alarm3"
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $xmlTemplate = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

    # Toast XML template
    [xml]$xmlTemplate = @"
    <toast scenario="reminder" useButtonStyle="true">
        <visual>
            <text placement="attribution">via $winTitle</text>
            <binding template="ToastGeneric" activationType="protocol">
                <image src="$($global:heroIMG)" placement="hero" />
                <image src="$($global:avatar)" placement="appLogoOverride" hint-crop="circle" />
                <group>
                <subgroup>
                    <text hint-align="left" hint-style="Header">$UserFirstName $UserLastName</text>
                    <text hint-style="Body" hint-wrap="true">$NotificationText</text>
                </subgroup>
                </group>
                <group>
                <subgroup>
                    <text placement="attribution">via $winTitle</text>
                </subgroup>
                </group>
            </binding>
        </visual>
        <actions>
            <action imageUri="$($global:openicon)" hint-toolTip="Text reply" hint-buttonStyle="Success" content="Open Webex" activationType="protocol" arguments="webexteams://im?email=$($global:email)" />
            <action imageUri="$($global:closeicon)" hint-buttonStyle="Critical" content="Dismiss" activationType="protocol" arguments="" />
        </actions>
        <audio src="$audSource" />
    </toast>
"@

    # Load
    $xmlToast = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xmlToast.LoadXml($xmlTemplate.OuterXml)

    # Display
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(" ").Show($xmlToast)
}

<# Starting script #>
function welcomeText {
    Write-Host "Welcome to Webex AvailaBOT !" -ForegroundColor Green
    $global:email = Read-Host -Prompt "Enter the user's e-mail you'd like to get a notification when available"
    Write-Host " "
    Write-Host " "
    Write-Host " "
    Write-Host " "
    Write-Host " "
    Write-Host " "
    Write-Host "The entered e-mail is" $global:email -ForegroundColor Green
    initialCheck
}

<# Fetch once Cisco's API to get user's data and current status #>
function initialCheck() {

    try {
        <# API call parameters #>
        $Header = @{
            "authorization" = "Bearer $token"
        }
        $Parameters = @{
            Method      = "GET"
            Uri         = "https://webexapis.com/v1/people?email=" + $global:email
            Headers     = $Header
            ContentType = "application/json; charset=utf-8"
            OutFile     = "./fetched-data/" + $ID + ".txt"
        }

        Invoke-RestMethod @Parameters
        $data = (Get-Content ./fetched-data/$ID.txt -Encoding UTF8) | ConvertFrom-Json
        $UserFirstName = $data.items.firstName
        $UserLastName = $data.items.lastName        
    }
    catch { 
        Write-Host "An error occured while fetching Webex's API" -ForegroundColor Red
        welcomeText 
    }

    <# Check if the user exist or not #>
    if ($UserFirstName) {
        Write-Host " "
        Write-Host("Fetching $UserFirstName $UserLastName's status") -ForegroundColor Green 
        loopCheck
    }
    else {
        Write-Host "This user doesn't exist. Please verify the e-mail" -ForegroundColor Red
        welcomeText
    }
}

<# Fetch Cisco's API each 30 seconds #>
function loopCheck() {
    do {
        try {
            <# API call parameters #>
            $Header = @{
                "authorization" = "Bearer $token"
            }
            $Parameters = @{
                Method      = "GET"
                Uri         = "https://webexapis.com/v1/people?email=" + $global:email
                Headers     = $Header
                ContentType = "application/json; charset=utf-8"
                OutFile     = "./fetched-data/" + $ID + ".txt"
            }
            <# Fetch user's data from Cisco API #>
            Invoke-RestMethod @Parameters
            $data = (Get-Content ./fetched-data/$ID.txt -Encoding UTF8) | ConvertFrom-Json
            $UserStatus = $data.items.status
            $LastSeen = Get-Date $data.items.lastActivity -UFormat "%A %d %B %Y à %T"      

            if ($UserStatus -ne "active") {
                <# Display the user's current status #>
                Write-Host($UserFirstName + " " + $UserLastName + " is " + $UserStatus + " (last activity: " + $LastSeen + ")") -ForegroundColor Green

                <# Create a progress bar to whow the remaining time before next check #>
                $seconds = 30
                1..$seconds |
                ForEach-Object { $percent = $_ * 100 / $seconds; 
                    Write-Progress -Activity "Fetching" -Status "Next check in $($seconds - $_) seconds..." -PercentComplete $percent;
                    Start-Sleep -Seconds 1
                }    
            } 
        }
        catch { Write-Host "An error occured while fetching Webex's API" -ForegroundColor Red }

    } while ($UserStatus -ne "active")

    <# Hide progress bar once the user is available #>
    Write-Progress -Completed -Activity "Fetching"

    <# Preparing the toast notification and closing the script window #>
    $NotificationText = ("is now available in Webex !")
    Write-Host ($UserFirstName + " " + $UserLastName + " is currently " + $UserStatus) -ForegroundColor Green

    <# Check if the user has an avatar and get it #>
    if ($data.items.avatar) {
        Invoke-WebRequest $data.items.avatar -OutFile "./fetched-data/$ID.png"
        $global:avatar = Get-ChildItem ./fetched-data/$ID.png
    }
    Show-Notification $NotificationText
    closeScript
}

<# Create a progress bar to whow the remaining time before closing windows #>
function closeScript() {
    Write-Host "This window will close automatically in 10 seconds..." -ForegroundColor Cyan
    $seconds = 10
    1..$seconds |
    ForEach-Object { $percent = $_ * 100 / $seconds; 
        Write-Progress -Activity "Closing" -Status "Closing window in $($seconds - $_) seconds..." -PercentComplete $percent;
        Start-Sleep -Seconds 1
    }

    <# Cleaning old files #>
    Write-Host "Temporary files are being cleaned..." -ForegroundColor Cyan
    if ($data.items.avatar) {
        Remove-Item ./fetched-data/$ID.png
    }
    Remove-Item ./fetched-data/$ID.txt
    Get-ChildItem –Path ./fetched-data/ –Recurse -include *.txt | Where-Object { $_.CreationTime -lt (Get-Date).AddMinutes(-5) } | Remove-Item
}

<# Start script #>
welcomeText