#Requires -module ImportExcel
Import-Module ImportExcel
#Main variables
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$configPath = "$scriptPath\Settings.xml"
[xml]$ScriptConfig = Get-Content $configPath

$logName = $ScriptConfig.Settings.PathSettings.LogFileName

$SecretFile = "$scriptPath\Secret.xml"

$CurrentDate = ((Get-date).ToUniversalTime()).ToString("yyyy-MM-dd")
$logPath = "$scriptPath\logs\$CurrentDate-$logName"

#Client secret and AAD app credentials
try {
    $Credentials = Import-Clixml -Path $SecretFile
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($Credentials.Password))
    $ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}
catch {
    break
}

#Email settings
$SenderAddress = $ScriptConfig.Settings.EmailSettings.Sender
$RecipientsAddresses = $ScriptConfig.Settings.EmailSettings.Recipients
$SMTPRelay = $ScriptConfig.Settings.EmailSettings.SMTPServer
$SMTPPort = $ScriptConfig.Settings.EmailSettings.EmailPort

if ($ClientSecret) {
    $AppID = $ScriptConfig.Settings.TenantSettings.appID
    $TenantId = $ScriptConfig.Settings.TenantSettings.tenantID
    $GraphUrl = $ScriptConfig.Settings.TenantSettings.GraphUrl
}else {
    break
}


#Functions
Function Get-GraphResult ($Url, $Token, $Method) {

    $Header = @{
        Authorization = "$($Token.token_type) $($Token.access_token)"
    }

    $PostSplat = @{
        ContentType = 'application/json'
        Method = $Method
        Header = $Header
        Uri = $Url
    }
    try {
        Invoke-RestMethod @PostSplat -ErrorAction Stop
    }
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        #break
    }
}

Function Get-GraphToken ($AppId, $AppSecret, $TenantID) {

    $AuthUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $Scope = "https://graph.microsoft.com/.default"

    $Body = @{
        client_id = $AppId
            client_secret = $AppSecret
            scope = $Scope
            grant_type = 'client_credentials'
    }

    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        # Create string by joining bodylist with '&'
        Body = $Body
        Uri = $AuthUrl
    }
    try {
        Invoke-RestMethod @PostSplat -ErrorAction Stop
    }
    catch {
        Write-Warning "Exception was caught: $($_.Exception.Message)" 
    }
}

function Get-Office365ActiveUsersDetail () {
    $graphApiVersion = "beta"
    $Resource = "reports/getOffice365ActiveUserDetail(period='D30')?`$format=application/json"
    $uri = "$graphUrl/$graphApiVersion/$($Resource)"
    $Method = "GET"
    try {
    
        $ResultResponse = Get-GraphResult -Url $uri -Token $Token -Method $Method
        $Result = $ResultResponse.value
        $ResultNextLink = $ResultResponse."@odata.nextLink"
        $page = $null
        if ($ResultNextLink) {
            while ($null -ne $ResultNextLink){
                $ResultResponse = Get-GraphResult -Url $ResultNextLink -Token $Token -Method $Method
                $ResultNextLink = $ResultResponse."@odata.nextLink"
                $Result += $ResultResponse.value
                $page += 1
                #Clear-Host
                Write-Verbose "Collecting data"
                Write-Verbose "Processing page number $page"
                #write-host $DevicesNextLink
            }
        }

        return $Result
        Write-Log -Message "Devices data successfuly collected"
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-log "Response content:`n$responseBody" -Level Error
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -leve Error
        Write-Log -Level Error -Message "Can't collect device info. Error message:`n$responseBody"
        break
    }
    return $Devices
}

function Get-SharePointSiteUsageDetail () {
    $graphApiVersion = "beta"
    $Resource = "reports/getSharePointSiteUsageDetail(period='D30')?`$format=application/json"
    $uri = "$graphUrl/$graphApiVersion/$($Resource)"
    $Method = "GET"
    try {
    
        $ResultResponse = Get-GraphResult -Url $uri -Token $Token -Method $Method
        $Result = $ResultResponse.value
        $ResultNextLink = $ResultResponse."@odata.nextLink"
        $page = $null
        if ($ResultNextLink) {
            while ($null -ne $ResultNextLink){
                $ResultResponse = Get-GraphResult -Url $ResultNextLink -Token $Token -Method $Method
                $ResultNextLink = $ResultResponse."@odata.nextLink"
                $Result += $ResultResponse.value
                $page += 1
                #Clear-Host
                Write-Verbose "Collecting data"
                Write-Verbose "Processing page number $page"
            }
        }

        return $Result
        Write-Log -Message "Devices data successfuly collected"
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-log "Response content:`n$responseBody" -Level Error
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -leve Error
        Write-Log -Level Error -Message "Can't collect device info. Error message:`n$responseBody"
        break
    }
    return $Devices
}

function Get-SkypeForBusinessActivityUserDetail () {
    $graphApiVersion = "beta"
    $Resource = "reports/getSkypeForBusinessActivityUserDetail(period='D30')?`$format=application/json"
    $uri = "$graphUrl/$graphApiVersion/$($Resource)"
    $Method = "GET"
    try {
    
        $ResultResponse = Get-GraphResult -Url $uri -Token $Token -Method $Method
        $Result = $ResultResponse.value
        $ResultNextLink = $ResultResponse."@odata.nextLink"
        $page = $null
        if ($ResultNextLink) {
            while ($null -ne $ResultNextLink){
                $ResultResponse = Get-GraphResult -Url $ResultNextLink -Token $Token -Method $Method
                $ResultNextLink = $ResultResponse."@odata.nextLink"
                $Result += $ResultResponse.value
                $page += 1
                #Clear-Host
                Write-Verbose "Collecting data"
                Write-Verbose "Processing page number $page"
            }
        }

        return $Result
        Write-Log -Message "Devices data successfuly collected"
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-log "Response content:`n$responseBody" -Level Error
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -leve Error
        Write-Log -Level Error -Message "Can't collect device info. Error message:`n$responseBody"
        break
    }
    return $Devices
}

function Get-TeamsUserActivityUserDetail () {
    $graphApiVersion = "beta"
    $Resource = "reports/getTeamsUserActivityUserDetail(period='D30')?`$format=application/json"
    $uri = "$graphUrl/$graphApiVersion/$($Resource)"
    $Method = "GET"
    try {
    
        $ResultResponse = Get-GraphResult -Url $uri -Token $Token -Method $Method
        $Result = $ResultResponse.value
        $ResultNextLink = $ResultResponse."@odata.nextLink"
        $page = $null
        if ($ResultNextLink) {
            while ($null -ne $ResultNextLink){
                $ResultResponse = Get-GraphResult -Url $ResultNextLink -Token $Token -Method $Method
                $ResultNextLink = $ResultResponse."@odata.nextLink"
                $Result += $ResultResponse.value
                $page += 1
                #Clear-Host
                Write-Verbose "Collecting data"
                Write-Verbose "Processing page number $page"
            }
        }

        return $Result
        Write-Log -Message "Devices data successfuly collected"
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-log "Response content:`n$responseBody" -Level Error
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -leve Error
        Write-Log -Level Error -Message "Can't collect device info. Error message:`n$responseBody"
        break
    }
    return $Devices
}

function Get-OneDriveUsageDetail () {
    $graphApiVersion = "beta"
    $Resource = "reports/getOneDriveUsageAccountDetail(period='D30')?`$format=application/json"
    $uri = "$graphUrl/$graphApiVersion/$($Resource)"
    $Method = "GET"
    try {
    
        $ResultResponse = Get-GraphResult -Url $uri -Token $Token -Method $Method
        $Result = $ResultResponse.value
        $ResultNextLink = $ResultResponse."@odata.nextLink"
        $page = $null
        if ($ResultNextLink) {
            while ($null -ne $ResultNextLink){
                $ResultResponse = Get-GraphResult -Url $ResultNextLink -Token $Token -Method $Method
                $ResultNextLink = $ResultResponse."@odata.nextLink"
                $Result += $ResultResponse.value
                $page += 1
                #Clear-Host
                Write-Verbose "Collecting data"
                Write-Verbose "Processing page number $page"
            }
        }

        return $Result
        Write-Log -Message "Devices data successfuly collected"
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-log "Response content:`n$responseBody" -Level Error
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -leve Error
        Write-Log -Level Error -Message "Can't collect device info. Error message:`n$responseBody"
        break
    }
    return $Devices
}

function Write-Log {

    [CmdletBinding()]
    Param (
        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [Alias('LogPath')]
        [string]$Path = $logPath,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoClobber
    
    )
    
    Begin {
        
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
    
    }
    Process {
        
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
                }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
                }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
                }
            }
        
        # Write log entry to $Path
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append
    }
    End {
    }
}

#Getting Graph Token
try {
    write-log -Message 'Getting Graph Token'
    $Token = Get-GraphToken -AppId $AppID -AppSecret $ClientSecret -TenantID $TenantId -ErrorAction Stop
    Write-Log "Token successfully issued"
}
catch {
    Write-Log -message "Can't get a token!" -Level Error
    break
}

#-----------------O365 Usage------------------#

Write-log -Message "Collecting O365Usage Data"
try {
    $Office365ActiveUsersDetail = Get-Office365ActiveUsersDetail -ErrorAction Stop
    Write-log -message "O365Usage data successfully collected"
}
catch {
    Write-log -Message "Can't collect O365Usage Data" -Level Error
}

if ($Office365ActiveUsersDetail) {
    $Office365ActiveUsersDetailReportPath = "$scriptPath\Reports\$CurrentDate-O365ActiveUserDetails.xlsx"

    $Office365ActiveUsersDetail | Select-Object @{L="Display Name";E={($_.displayName)}}, `
    @{L="User Principal Name";E={($_.'userPrincipalName').tolower()}}, @{L="Assigned Licenses";E={$_.AssignedProducts -join ", "}}, `
    @{L="Has Exchange License";E={$_.hasExchangeLicense}}, @{L="Exchange Last SignIn";E={$_.exchangeLicenseAssignDate}}, `
    @{L="Has OneDrive License'";E={$_.hasOneDriveLicense}}, @{L="OneDrive Last Activity Date";E={$_.oneDriveLastActivityDate}},
    @{L="Has SharePoint License";E={$_.hasSharePointLicense}}, @{L="SharePoint Last Activity Date";E={$_.sharePointLastActivityDate}}, `
    @{L="Has Teams License";E={$_.hasTeamsLicense}}, @{L="Teams Last Activity Date";E={$_.teamsLastActivityDate}}, `
    @{L="Has Skype For Business License";E={$_.hasSkypeForBusinessLicense}}, @{L="Skype For Business Last Activity Date";E={$_.skypeForBusinessLastActivityDate}},
    @{L="Are Services Used";E={`
        if((($_.exchangeLicenseAssignDate) `
        -or ($_.oneDriveLastActivityDate) `
        -or ($_.sharePointLastActivityDate) `
        -or ($_.teamsLastActivityDate) `
        -or ($_.skypeForBusinessLastActivityDate)) `
        -and ($_.assignedProducts))  {$true} else  {$false}}} | `
    Export-Excel -Path $Office365ActiveUsersDetailReportPath -AutoSize -AutoFilter -TableStyle Medium2
    write-log "O365Usage has been exported here: $Office365ActiveUsersDetailReportPath"
}else {
    Write-log -Level Error -Message "O365UsageReport was not exported"
}

#----------------SPO Site Usage Details-------------------#

Write-log -Message "Collecting SPOSiteUsage Data"
try {
    $SharePointSiteUsageDetail = Get-SharePointSiteUsageDetail -ErrorAction Stop
    Write-log -message "SPOSiteUsage data successfully collected"
}
catch {
    Write-log -Message "Can't collect SPOSiteUsage Data" -Level Error
}

if ($SharePointSiteUsageDetail) {
    $SharePointSiteUsageDetailReportPath = "$scriptPath\Reports\$CurrentDate-SPOSiteUsageDetails.xlsx"

    $SharePointSiteUsageDetail | Select-Object @{L="Site ID";E={($_.siteId)}}, `
    @{L='Site URL'; E={$_.siteUrl}}, @{L='Owner/Group Display Name'; E={$_.ownerDisplayName}}, @{L='Owner or Group UPN'; E={$_.ownerPrincipalName}}, `
    @{L='Is Site Deleted'; E={$_.isDeleted}}, @{L='Site Last Activity Date'; E={$_.isDeleted}}, @{L='File Count'; E={$_.fileCount}}, `
    @{L='Active File Count'; E={$_.activeFileCount}}, @{L='Page View Count'; E={$_.pageViewCount}}, @{L='Visited Page Count'; E={$_.visitedPageCount}}, `
    @{L='Storage Used (MB)'; E={$([math]::round($_.storageUsedInBytes/1MB, 2))}}, @{L='Storage Allocated (GB)'; E={$([math]::round($_.storageAllocatedInBytes/1Gb, 2))}}, `
    @{L='Site Template'; E={$_.rootWebTemplate}} | `
    Export-Excel -Path $SharePointSiteUsageDetailReportPath -AutoSize -AutoFilter -TableStyle Medium2
    write-log "SPOSiteUsage has been exported here: $SharePointSiteUsageDetailReportPath"
}else {
    Write-log -Level Error -Message "SPOSiteUsage was not exported"
}

#----------------SfB Activity User Detail------------------#

Write-log -Message "Collecting SfBActivityUserDetail Data"
try {
    $SkypeForBusinessActivityUserDetail = Get-SkypeForBusinessActivityUserDetail -ErrorAction Stop
    Write-log -message "SfBActivityUserDetail data successfully collected"
}
catch {
    Write-log -Message "Can't collect SfBActivityUserDetail Data" -Level Error
}

if ($SkypeForBusinessActivityUserDetail) {
    $SkypeForBusinessActivityUserDetailReportPath = "$scriptPath\Reports\$CurrentDate-SfBActivityUserDetail.xlsx"

    $SkypeForBusinessActivityUserDetail | Select-Object @{L='User UPN'; E={$_.userPrincipalName}}, @{L='Deleted'; E={$_.isDeleted}}, `
    @{L='Deleted Date'; E={$_.deletedDate}}, @{L='Last Activity Date'; E={$_.lastActivityDate}}, @{L="Assigned Licenses";E={$_.AssignedProducts -join ", "}}, `
    @{L='P2P-Total'; E={$_.totalPeerToPeerSessionCount}}, @{L='Organized Conf.-Total'; E={$_.totalOrganizedConferenceCount}}, `
    @{L='Participated Conf. - Total'; E={$_.totalParticipatedConferenceCount}}, @{L='P2P-Last'; E={$_.peerToPeerLastActivityDate}}, `
    @{L='Last Org-ed Conf.'; E={$_.organizedConferenceLastActivityDate}}, @{L='Last Particip. Conf.'; E={$_.participatedConferenceLastActivityDate}}, `
    @{L='P2P-IM'; E={$_.peerToPeerIMCount}}, @{L='P2P-Audio'; E={$_.peerToPeerAudioCount}}, @{L='P2P-Audio(minutes)'; E={$_.peerToPeerAudioMinutes}}, `
    @{L='P2P-Video'; E={$_.peerToPeerVideoCount}}, @{L='P2P-Video(minutes)'; E={$_.peerToPeerVideoMinutes}}, @{L='P2P-App Sharing'; E={$_.peerToPeerAppSharingCount}}, `
    @{L='P2P-File Transfer'; E={$_.peerToPeerFileTransferCount}}, @{L='IM Organized Conf.'; E={$_.organizedConferenceIMCount}}, `
    @{L='Audio Video-Organized Conf.'; E={$_.organizedConferenceAudioVideoCount}}, @{L='Audio Video-Organized(minutes)'; E={$_.organizedConferenceAudioVideoMinutes}}, `
    @{L='App Sharing-Organized Conf.'; E={$_.organizedConferenceAppSharingCount}} ,@{L='Organized Conf. Web'; E={$_.organizedConferenceWebCount}}, `
    @{L='Organized Conf. Dial-In/Out 3rdParty'; E={$_.organizedConferenceDialInOut3rdPartyCount}}, @{L='Organized Conf. CloudDial-In/Out MS'; E={$_.organizedConferenceCloudDialInOutMicrosoftCount}}, `
    @{L='Organized Conf. CloudDial-In MS(minutes)'; E={$_.organizedConferenceCloudDialInMicrosoftMinutes}}, @{L='Organized Conf. CloudDial-Out MS(minutes)'; E={$_.organizedConferenceCloudDialOutMicrosoftMinutes}}, `
    @{L='Par-ed Conf. IM'; E={$_.participatedConferenceIMCount}}, @{L='Par-ed Conf. AudioVideo'; E={$_.participatedConferenceAudioVideoCount}}, `
    @{L='Par-ed Conf. AudioVideo Minute'; E={$_.participatedConferenceAudioVideoMinutes}}, @{L='Par-ed Conf. AppSharing'; E={$_.participatedConferenceAppSharingCount}}, `
    @{L='Par-ed Conf. Web'; E={$_.participatedConferenceWebCount}}, @{L='Par-ed Conf. DialInOut 3rdParty'; E={$_.participatedConferenceDialInOut3rdPartyCount}}, `
    @{L='ReportRefreshDate'; E={$_.reportRefreshDate}}, @{L='Report Period'; E={$_.reportPeriod}} | `
    Export-Excel -Path $SkypeForBusinessActivityUserDetailReportPath -AutoSize -AutoFilter -TableStyle Medium2
    write-log "SfBActivityUserDetail has been exported here: $SkypeForBusinessActivityUserDetailReportPath"
}else {
    Write-log -Level Error -Message "SfBActivityUserDetail was not exported"
}


#----------------Teams Activity detail------------------#

Write-log -Message "Collecting TeamsActivity Data"
try {
    $TeamsUserActivityUserDetail = Get-TeamsUserActivityUserDetail -ErrorAction Stop
    Write-log -message "TeamsActivity data successfully collected"
}
catch {
    Write-log -Message "Can't collect TeamsActivity Data" -Level Error
}

if ($TeamsUserActivityUserDetail) {
    $TeamsUserActivityUserDetailReportPath = "$scriptPath\Reports\$CurrentDate-TeamsUserActivityUserDetail.xlsx"

    $TeamsUserActivityUserDetail | Select-Object @{L='Report Date'; E={$_.reportRefreshDate}}, @{L='User UPN'; E={$_.userPrincipalName}}, @{L='Deleted'; E={$_.isDeleted}}, `
    @{L='Deleted Date'; E={$_.deletedDate}}, @{L='Last Activity'; E={$_.lastActivityDate}}, @{L="Assigned Licenses";E={$_.AssignedProducts -join ", "}}, `
    @{L='Team Chat Messages'; E={$_.teamChatMessageCount}}, @{L='Private Chat Messages'; E={$_.privateChatMessageCount}},  @{L='Calls'; E={$_.callCount}}, `
    @{L='Meetings'; E={$_.meetingCount}} | `
    Export-Excel -Path $TeamsUserActivityUserDetailReportPath -AutoSize -AutoFilter -TableStyle Medium2
    write-log "TeamsActivity has been exported here: $TeamsUserActivityUserDetailReportPath"
}else {
    Write-log -Level Error -Message "TeamsActivity was not exported"
}

#----------------OneDrive Usage------------------#

Write-log -Message "Collecting OneDriveUsageDetail Data"
try {
    $OneDriveUsageDetail = Get-OneDriveUsageDetail -ErrorAction Stop
    Write-log -message "OneDriveUsageDetail data successfully collected"
}
catch {
    Write-log -Message "Can't collect OneDriveUsageDetail Data" -Level Error
}

if ($OneDriveUsageDetail) {
    $OneDriveUsageDetailReportPath = "$scriptPath\Reports\$CurrentDate-OneDriveUsageDetail.xlsx"

    $OneDriveUsageDetail | Select-Object @{L='Report Date'; E={$_.reportRefreshDate}}, `
    @{L='Site URL'; E={$_.siteUrl}}, @{L='Owner Display Name'; E={$_.ownerDisplayName}}, `
    @{L='Owner Principal Name'; E={$_.ownerPrincipalName}}, @{L='Deleted'; E={$_.isDeleted}}, @{L='Deleted Date'; E={$_.deletedDate}}, `
    @{L='Last Activity'; E={$_.lastActivityDate}}, @{L='Files'; E={$_.fileCount}}, `
    @{L='Active Files'; E={$_.activeFileCount}}, @{L='Storage Used (MB)'; E={$([math]::round($_.storageUsedInBytes/1MB, 2))}}, `
    @{L='Storage Allocated (GB)'; E={$([math]::round($_.storageAllocatedInBytes/1Gb, 2))}} | `
    Export-Excel -Path $OneDriveUsageDetailReportPath -AutoSize -AutoFilter -TableStyle Medium2
    write-log "OneDriveUsageDetail has been exported here: $OneDriveUsageDetailReportPath"
}else {
    Write-log -Level Error -Message "OneDriveUsageDetail was not exported"
}

$Attachment = @()

if ([System.IO.File]::Exists($Office365ActiveUsersDetailReportPath)) {
    $Attachment += $Office365ActiveUsersDetailReportPath
}
if ([System.IO.File]::Exists($SharePointSiteUsageDetailReportPath)) {
    $Attachment += $SharePointSiteUsageDetailReportPath
}
if ([System.IO.File]::Exists($SkypeForBusinessActivityUserDetailReportPath)) {
    $Attachment += $SkypeForBusinessActivityUserDetailReportPath
}
if ([System.IO.File]::Exists($TeamsUserActivityUserDetailReportPath)) {
    $Attachment += $TeamsUserActivityUserDetailReportPath
}
if ([System.IO.File]::Exists($OneDriveUsageDetailReportPath)) {
    $Attachment += $OneDriveUsageDetailReportPath
}

try {
    Send-MailMessage `
        -to $RecipientsAddresses `
        -Subject "O365 Services Reports on $((get-date).ToString("dd/MM/yyy"))" `
        -Body "Hello Team! Please see O365 Services Reports $((get-date).ToString("dd/MM/yyy")) <br />" `
        -from $SenderAddress `
        -Attachments $Attachment `
        -SmtpServer $SMTPRelay -BodyAsHtml -Port $SMTPPort
        Write-Log -Message "Email has been sent" 
}
catch {
        Write-Log -Message "$($_.Exception.Message)" -Path -Level Error  
        break
}
