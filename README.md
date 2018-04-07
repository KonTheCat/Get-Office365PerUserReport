# Get-Office365PerUserReport
PowerShell script to produce a per-user report from Office365. Was written mostly to create a modular framework for addressing these kinds of requests in a standard way.

Usage - dot-source into the console. The function Get-Office365PerUserReport will handle authentication into Office365, MSOL, and Azure AD. It will also remove the PS remoting session into Office365 and will disconnect from Azure AD after the report is executed. 

To add functionality - add functions that will accept the UserPrincipalName as identifying the user above the final function in the script. Add lines of $ReportObject | Add-Member accordinly.  

SYNOPSIS

Gets a report of Office365 users that have licences. Puts it on the desktop.

EXAMPLE

Get-Office365Report
