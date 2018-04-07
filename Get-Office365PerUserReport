function Get-Office365ReportLicences {
    [CmdletBinding()]
    Param($Identity)

    $LicenceTable = @{
        "O365_BUSINESS_ESSENTIALS" = "Office 365 Business Essentials"
        "ENTERPRISEPACK" = "Office 365 Enterprise E3"
        "O365_BUSINESS_PREMIUM" = "Office 365 Business Premium"
        "O365_BUSINESS" = "Office 365 Business"
        "ENTERPRISEPREMIUM" = "Office 365 Enterprise E5"
        "LITEPACK_P2" = "Office 365 Small Business Premium"
        "MIDSIZEPACK" = "Office 365 Midsize Business"
        "ENTERPRISEPACKWSCAL" = "Office 365 Enterprise E4"
        "STANDARDPACK_STUDENT" =  "Microsoft Office 365 (Plan A1) for Students"
        "STANDARDPACK_FACULTY" =  "Microsoft Office 365 (Plan A1) for Faculty"
        "STANDARDWOFFPACK_FACULTY" =  "Office 365 Education E1 for Faculty"
        "STANDARDWOFFPACK_STUDENT" =  "Microsoft Office 365 (Plan A2) for Students"
        "STANDARDWOFFPACK_IW_STUDENT" =  "Office 365 Education for Students"
        "STANDARDWOFFPACK_IW_FACULTY" =  "Office 365 Education for Faculty"
        "EOP_ENTERPRISE_FACULTY" =  "Exchange Online Protection for Faculty"
        "EXCHANGESTANDARD_STUDENT" =  "Exchange Online (Plan 1) for Students"
        "OFFICESUBSCRIPTION_STUDENT" =  "Office ProPlus Student Benefit"
        "STANDARDPACK_GOV" =  "Microsoft Office 365 (Plan G1) for Government"
        "STANDARDWOFFPACK_GOV" =  "Microsoft Office 365 (Plan G2) for Government"
        "ENTERPRISEPACK_GOV" =  "Microsoft Office 365 (Plan G3) for Government"
        "ENTERPRISEWITHSCAL_GOV" =  "Microsoft Office 365 (Plan G4) for Government"
        "DESKLESSPACK_GOV" =  "Microsoft Office 365 (Plan K1) for Government"
        "ESKLESSWOFFPACK_GOV" =  "Microsoft Office 365 (Plan K2) for Government"
        "EXCHANGESTANDARD_GOV" =  "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
        "EXCHANGEENTERPRISE_GOV" =  "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
        "SHAREPOINTDESKLESS_GOV" =  "SharePoint Online Kiosk"
        "EXCHANGE_S_DESKLESS_GOV" =  "Exchange Kiosk"
        "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
        "RMS_S_ENTERPRISE_GOV" =  "Windows Azure Active Directory Rights Management"
        "OFFICESUBSCRIPTION_GOV" =  "Office ProPlus"
        "MCOSTANDARD_GOV" =  "Lync Plan 2G"
        "SHAREPOINTWAC_GOV" =  "Office Online for Government"
        "SHAREPOINTENTERPRISE_GOV" =  "SharePoint Plan 2G"
        "EXCHANGE_S_ENTERPRISE_GOV" =  "Exchange Plan 2G"
        "EXCHANGE_S_ARCHIVE_ADDON_GOV" =  "Exchange Online Archiving"
        "LITEPACK" =  "Office 365 (Plan P1)"
        "STANDARDPACK" =  "Microsoft Office 365 (Plan E1)"
        "STANDARDWOFFPACK" =  "Microsoft Office 365 (Plan E2)"
        "DESKLESSPACK" =  "Office 365 (Plan K1)"
        "EXCHANGEARCHIVE" =  "Exchange Online Archiving"
        "EXCHANGETELCO" =  "Exchange Online POP"
        "SHAREPOINTSTORAGE" =  "SharePoint Online Storage"
        "SHAREPOINTPARTNER" =  "SharePoint Online Partner Access"
        "PROJECTONLINE_PLAN_1" =  "Project Online (Plan 1)"
        "PROJECTONLINE_PLAN_2" =  "Project Online (Plan 2)"
        "PROJECT_CLIENT_SUBSCRIPTION" =  "Project Pro for Office 365"
        "VISIO_CLIENT_SUBSCRIPTION" =  "Visio Pro for Office 365"
        "INTUNE_A" =  "Intune for Office 365"
        "CRMSTANDARD" =  "CRM Online"
        "CRMTESTINSTANCE" =  "CRM Test Instance"
        "ONEDRIVESTANDARD" =  "OneDrive"
        "WACONEDRIVESTANDARD" =  "OneDrive Pack"
        "SQL_IS_SSIM" =  "Power BI Information Services"
        "BI_AZURE_P1" =  "Power BI Reporting and Analytics"
        "EOP_ENTERPRISE" =  "Exchange Online Protection"
        "PROJECT_ESSENTIALS" =  "Project Lite"
        "CRMIUR" =  "CRM for Partners"
        "NBPROFESSIONALFORCRM" =  "Microsoft Social Listening Professional"
        "AAD_PREMIUM" =  "Azure Active Directory Premium"
        "MFA_PREMIUM" =  "Azure Multi-Factor Authentication"
        "DESKLESSPACK_YAMMER" =  "Office 365 Enterprise K1 with Yammer"
        "DESKLESSWOFFPACK" =  "Office 365 Enterprise K2"
        "EXCHANGE_L_STANDARD" =  "Exchange Online (Plan 1)"
        "EXCHANGE_S_DESKLESS" =  "Exchange Online Kiosk"
        "EXCHANGE_S_STANDARD" =  "Exchange Online (Plan 2)"
        "EXCHANGE_S_STANDARD_MIDMARKET" =  "Exchange Online (Plan 1)"
        "MCOLITE" =  "Lync Online (Plan 1)"
        "MCOSTANDARD" =  "Lync Online (Plan 2)"
        "MCOSTANDARD_MIDMARKET" =  "Lync Online (Plan 1)"
        "MCVOICECONF" =  "Lync Online (Plan 3)"
        "OFFICESUBSCRIPTION" =  "Office ProPlus"
        "RMS_S_ENTERPRISE" =  "Azure Active Directory Rights Management"
        "SHAREPOINTDESKLESS" =  "SharePoint Online Kiosk"
        "SHAREPOINTENTERPRISE" =  "SharePoint Online (Plan 2)"
        "SHAREPOINTENTERPRISE_MIDMARKET" =  "SharePoint Online (Plan 1)"
        "SHAREPOINTLITE" =  "SharePoint Online (Plan 1)"
        "SHAREPOINTWAC" =  "Office Online"
        "YAMMER_ENTERPRISE" =  "Yammer"
        "YAMMER_MIDSIZE" =  "Yammer"
        "EXCHANGEENTERPRISE" = "Exchange Online (2)"
    }

    $lic = Get-MsolUser -UserPrincipalName $Identity | Select-Object -Property Licenses
    $lic = $lic.Licenses.AccountSku | Select-Object -ExpandProperty SkuPartNumber

    $friendlylics = @()

    Foreach ($l in $lic) {
        $outlic = $LicenceTable.$l
        if ($outlic) {
            $friendlylics += $outlic
        } else {
            $friendlylics += $l
        }
    }
    Return $friendlylics
}

function Get-Office365ReportPrimaryEmailAddress {
    [CmdletBinding()]
    Param($Identity)

    $primaryaddress = Get-MSOLUser -UserPrincipalName $Identity | Select-Object -ExpandProperty ProxyAddresses | Where-Object {$_ -cmatch '^SMTP\:.*'}
    $primaryaddress = ($primaryaddress -split ":")[1]
    Return $primaryaddress 
}  

function Get-Office365ReportAllEmailAddresses {

    [CmdletBinding()]
    Param($Identity)

    $AllEmails = Get-MSOLUser -UserPrincipalName $Identity | Select-Object -ExpandProperty ProxyAddresses
    $outAllEmails = @()
    foreach ($email in $AllEmails) {
        $outemail = ($email -split ":")[1]

        #because the onmicrosoft email address is not needed for reporting 
        if ($outemail -notlike "*.onmicrosoft.com") {
            $outAllEmails += $outemail
        }
        
    }

    Return $outAllEmails
    
}

function Get-Office365ReportMailGroupMembership {
    [CmdletBinding()]
    Param($Identity)

    $MailGroupMemberOf = Get-AzureADUser -SearchString $Identity | Get-AzureADUserMembership | Where-Object {$_.MailEnabled -eq "True"} | Select-Object -ExpandProperty DisplayName
    Return $MailGroupMemberOf
    
}

function Get-Office365ReportMailboxesMemberOf {
    [CmdletBinding()]
    Param($Identity,$PermissionsTable)

    $mailboxes = $PermissionsTable | Where-Object {$_.user -eq $Identity} | Select-Object -ExpandProperty Identity

    $outMailboxesByEmail = @()
    foreach ($mailbox in $mailboxes) {
        $outMailbox = Get-mailbox -Identity $mailbox | Select-Object -ExpandProperty PrimarySmtpAddress
        $outMailboxesByEmail += $outMailbox
    }
    Return $outMailboxesByEmail
}

function Get-Office365PerUserReport {

        <#
    .SYNOPSIS
        Gets a report of Office365 users that have licences. Puts it on the desktop.    
    .EXAMPLE
        Get-Office365Report
    #>

    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$false)]
    [string]$ReportPath
    )

    #handle auth
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $Session -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
    Connect-MsolService -Credential $UserCredential
    Connect-AzureAD -Credential $UserCredential

    #handle getting users - we really only care about getting licenced users, regardless of enabled or disabled status, to catch things 
    Write-Output "Currently getting all the users. This can take a little while." 
    $users = Get-MsolUser -all | Where-Object {$_.Licenses -ne $null} | Select-Object -ExpandProperty userprincipalname
    #this is so we dont get-mailbox every time 
    Write-Output "Currently getting all the mailbox permissions. This can take longer than a little while." 
    $permissions = Get-Mailbox | Get-MailboxPermission 

    $Report = @() 
    #handle for each user
    foreach ($user in $users)
    {
    $ReportObject = New-Object PSobject
    Write-Output "Currently running the report for $user." 
    $ReportObject | Add-Member -type NoteProperty -name 'Name' -value (Get-MsolUser -UserPrincipalName $user | Select-Object -ExpandProperty Displayname)
    $ReportObject | Add-Member -type NoteProperty -name 'Primary Email Address' -value (Get-Office365ReportPrimaryEmailAddress -Identity $user)
    $ReportObject | Add-Member -type NoteProperty -name 'Email Addresses' -value ((Get-Office365ReportAllEmailAddresses -Identity $user) -join "; ")
    $ReportObject | Add-Member -type NoteProperty -name 'Licences' -value ((Get-Office365ReportLicences -Identity $user) -join "; ")
    $ReportObject | Add-Member -type NoteProperty -name 'Group Membership' -value ((Get-Office365ReportMailGroupMembership -Identity $user) -join "; ")
    $ReportObject | Add-Member -type NoteProperty -name 'Mailbox Access' -value ((Get-Office365ReportMailboxesMemberOf -Identity $user -PermissionsTable $permissions) -join "; ")
    $Report += $ReportObject
    }

    #handle report output
    $ReportName = ((Get-MsolCompanyInformation).DisplayName) + ' Office 365 Report ' + (Get-Date -Format "yyyy-MM-dd-HH-mm") + '.csv'
    if ($ReportPath) {
        $Report | Export-Csv -Path $ReportPath -NoTypeInformation
    } else {
        $ReportPath = $env:USERPROFILE + '\Desktop\' + $ReportName
        $Report | Export-Csv -Path $ReportPath -NoTypeInformation
    }

    #handle cleanup
    Remove-PSSession $Session
    Disconnect-AzureAD
    Write-Warning "Please exit the console to complete cleanup of connections!"
}
