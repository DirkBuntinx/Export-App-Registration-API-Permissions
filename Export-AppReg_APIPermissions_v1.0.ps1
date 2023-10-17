#################################################################################################################################
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. # 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  #
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.               #
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  #
# code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software   #
# product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the  #
# Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims   #
# or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.                 #
#################################################################################################################################
#----------------------------------------------------------------------              
#-     DO NOT CHANGE ANY CODE BELOW THIS LINE                         -
#----------------------------------------------------------------------
#-                                                                    -
#-                           Author:  Dirk Buntinx                    -
#-                           Date:    26/9/2023                       -
#-                           Version: v1.0                            -
#-                                                                    -
#----------------------------------------------------------------------

<#
.SYNOPSIS
Script to export the API Permissions for each Application Registration in Azure AD.
IMPORTANT: This is a READ ONLY script; it will only read information for AzureAD and make no modifications.

.DESCRIPTION
This script will export the API Permissions for each Application Registration in Azure AD.

Special thanks to my colleagues for providing valuable feedback and validating the script’s functionality:
-	Dan Bagley
-	Smart Kamolratanapiboon
-	Brij Raj Singh
-	Angélique Conde
-	Max Vaughn
-	Bonny Eapen
-	David Castro Koschny

The script uses 3 mandatory parameters:
---------------------------------------

1) UsePSModule: This parameter determines which PowerShell Module will be used to connect to Azure AD. Two possible values that can be autocomplete by using the <Tab> key
                    
    - Microsoft.Graph: this option will force the script to use 'Microsoft.Graph' module minimum version 2.6.1 to connect to Azure AD.
                        The script will validate if the correct version is installed and uses the Graph API Permission 'Application.Read.All' to run
                        the Graph cmdlet's Connect-MgGraph, Get-MgApplication and Get-MgServicePrincipal. 
                        Make sure that an AzureAD Administrator has correctly setup the 'Microsoft Graph Command Line Tools' Enterprise Application registration by granting Admin Consent 
                        for the API Permission 'Application.Read.All', if not, you will be prompted to provide Admin Consent at script runtime.

                        IMPORTANT: To install the required 'Microsoft.Graph' PowerShell module, please run the following cmdlet: 
                        install-module -Name Microsoft.Graph -MinimumVersion '2.6.1' -Repository:PSGallery

    - AzureAD: this option will force the script to use 'AzureAD' PowerShell module minimum version 2.0.2.182 to connect to Azure AD.
                The script will validate if the correct version is installed and uses the AzureAD cmdlet's Connect-AzureAD, Get-AzureADApplication and Get-AzureADServicePrincipal
                IMPORTANT: Please note that the AzureAD PowerShell module is scheduled to be retired on March 30, 2024
                More information see articles: 
                    - https://techcommunity.microsoft.com/t5/microsoft-entra-azure-ad-blog/important-azure-ad-graph-retirement-and-powershell-module/ba-p/3848270#:~:text=One%20year%20ago%20we%20communicated,deprecated%20on%20June%2030%2C%202023.
                    - https://learn.microsoft.com/en-us/powershell/microsoftgraph/migration-steps?view=graph-powershell-1.0

                IMPORTANT: To install the required 'AzureAD' PowerShell module, please run the following cmdlet: 
                install-module -Name AzureAD -MinimumVersion '2.0.2.182' -Repository:PSGallery

2) ExportAPIPermissions: This parameter determines which Azure AD API Permissions to export, this parameter can have the following values:

    - All: This will export ALL the 'Delegated' and 'Application' API permissions API Permissions for each Application Registration in Azure AD.
        IMPORTANT: Take extra care of using the "All" value for Switch -ExportAPIPermissions as this will process every API Permission for ALL Resource Access groups for every Application 
        registration in AzureAD, depending on the number of Application registrations in your tenant, this process might take a very long time.
        Example, this will export API permissions for Resource Access Groups like PowerBI, Intune,Dynamics, all Azure related, etc

    - OutlookRESTv2: This will export both the 'Delegated' and 'Application' API permissions used by the Outlook REST API for each Application Registration in Azure AD.

    - EWS: This will export both the 'Delegated' and 'Application' API permissions used by the 'Exchange Web Services (EWS)' API for each Application Registration in Azure AD.

    - Graph: This will export both the 'Delegated' and 'Application' API permissions for the 'Microsoft Graph' API for each Application Registration in Azure AD.

    - O365MgmtAPI: This will export both the 'Delegated' and 'Application' API permissions for the 'Office 365 Management APIs' API for each Application Registration in Azure AD.

    - POP: This will export both the 'Delegated' and 'Application' API permissions used by the 'EPOP' API for each Application Registration in Azure AD.

    - IMAP: This will export both the 'Delegated' and 'Application' API permissions used by the 'IMAP' API for each Application Registration in Azure AD.

    - SMTP: This will export both the 'Delegated' and 'Application' API permissions used by the 'SMTP' API for each Application Registration in Azure AD.

    - ReportingServices: This will export both the 'Delegated' and 'Application' API permissions used by the 'Reporting Services' API for each Application Registration in Azure AD.

    - ExchangePowerShell: This will export both the 'Delegated' and 'Application' API permissions used by the 'ExchangePowershell' API for each Application Registration in Azure AD.


3) OutputPath: Define the Path to the directory where the output file will be saved. All selected data will be exported to a Tab separated csv file 'Export-AppReg_APIPermissions_<timestamp>.csv'
                The export file will be a Tab Separated CSV file as the Application Registration Name allows for comma's (',').
                In order to correctly view the files content, the recommendation is to import the data into Excel by using the 'Data Import' method:
                    - Open a Blank Workbook in Excel
                    - Go to the "Data" Tab
                    - Select "Get Data" and select "From File" and click "From Text/csv" and follow the prompts to import the data.

.EXAMPLE
.\Export-AppReg_APIPermissions_v0.6.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:EWS -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'EWS' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-AppReg_APIPermissions_v0.6.ps1 -UsePSModule:AzureAD -ExportAPIPermissions:EWS -OutputPath:'C:\temp'
The will cause PowerShell to use the 'AzureAD' module to retrieve all the Application registrations from AzureAD and will export ALL the 'EWS' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-AppReg_APIPermissions_v0.6.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:OutlookRESTv2 -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'Graph' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-AppReg_APIPermissions_v0.6.ps1 -UsePSModule:AzureAD -ExportAPIPermissions:OutlookRESTv2 -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'Graph' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-AppReg_APIPermissions_v0.6.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:All -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL API Permissions for ALL Resource Access Group 
to a file called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory
IMPORTANT: Take extra care of using the "All" value for Switch -ExportAPIPermissions as this will process every API Permission for ALL Resource Access groups for every Application registratioon in AzureAD, depending on the
number of Application registrations in your tenant, this process might take a very long time.
Example, this will export API permissions for Resource Access Groups like PowerBI, Intune,Dynamics, all Azure related, etc

#>

[CmdletBinding(DefaultParameterSetName ="Default")]
param(

    [Parameter(ParameterSetName="Default",Mandatory=$true, Position=0, HelpMessage="Specify which PowerShell module to use to connect to AzureAD, choices are 'Microsoft.Graph' module and 'AzureAD' module")]
    [ValidateSet('Microsoft.Graph','AzureAD')]
    [string]$UsePSModule,

    [Parameter(ParameterSetName="Default",Mandatory=$true, Position=1, HelpMessage="Switch to specifify which API Permisions to export for esach Application registration")]
    [ValidateSet('All','OutlookRESTv2','EWS', 'Graph', 'O365MgmtAPI', 'POP', 'IMAP', 'SMTP', 'ReportingServices', 'ExchangePowerShell')]
    [string]$ExportAPIPermissions,

    [Parameter(ParameterSetName="Default", Mandatory=$true, Position=2, HelpMessage="Define the Path to the directory where the output file will be saved.")]
    [string]$OutputPath=($(Get-Location).Path)
)  

###################################
# Declaring Script wide Variables #
###################################

$Date = [DateTime]::Now
$Script:StartTime = '{0:MM/dd/yyyy HH:mm:ss}' -f $Date
$Script:FileName = "Export-AppReg_APIPermissions_$('{0:MMddyyyyHHmms}' -f $Date).csv"
$Script:OutputStream = $null
$Script:Tab = [char]9
$Script:csvOutput = ""
$Script:csvAppInfo = ""
$Script:csvResourceInfo = ""
$Script:csvPermissionInfo = ""
$Script:AzureAppRegs = $null

# Define a Hash table to cache the ServicePrincipal objects for faster lookup
$Script:CacheServicePrincipal=@{}

# Define the ResourceAccessID for 'Office 365 Exchange Online' 00000002-0000-0ff1-ce00-000000000000
[string]$Script:Office365ExchangeOnline = "00000002-0000-0ff1-ce00-000000000000"

# Define the ResourceAccessID for 'Microsoft Graph'  00000003-0000-0000-c000-000000000000
[string]$Script:MicrosoftGraph = "00000003-0000-0000-c000-000000000000"

# Define the ResourceAccessID for 'Office 365 Management APIs' c5393580-f805-4401-95e8-94b7a6ef2fc2
[string]$Script:O365ManagementAPI = "c5393580-f805-4401-95e8-94b7a6ef2fc2"

# Define Array containing all Outlook REST API Delegated Permissions
$Script:Delegated_OutlookRESTPermissions = @("PeopleSettings.Read.All", "PeopleSettings.ReadWrite.All", "ReportingWebService.Read", "Organization.ReadWrite.All", 
    "Organization.Read.All", "Mail.ReadBasic", "Notes.Read", "Notes.ReadWrite", "User.Read.All", "User.ReadBasic.All", "MailboxSettings.Read", "Calendars.Read.Shared", 
    "Calendars.ReadWrite.Shared", "Mail.Send.Shared", "Mail.ReadWrite.Shared", "Mail.Read.Shared", "Contacts.ReadWrite.Shared", "Contacts.Read.Shared", "Tasks.Read.Shared", 
    "Tasks.ReadWrite.Shared", "Mail.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.Read", "Calendars.ReadWrite", "Contacts.Read", "Contacts.ReadWrite", "Group.Read.All", 
    "Group.ReadWrite.All", "User.Read", "User.ReadWrite", "User.ReadBasic.All", "People.Read", "People.ReadWrite", "Tasks.Read", "Tasks.ReadWrite", "MailboxSettings.ReadWrite", 
    "Contacts.ReadWrite.All", "Contacts.Read.All", "Calendars.ReadWrite.All", "Calendars.Read.All", "Mail.Send.All", "Mail.ReadWrite.All", "Mail.Read.All", "Place.Read.All", 
    "OPX.MyDay", "OPX.MyDay.Shared", "OPX.MyDay.All")


# Define Array containing all Outlook REST API Application Permissions
$Script:Application_OutlookRESTPermissions = @("PeopleSettings.ReadWrite.All", "PeopleSettings.Read.All", "Organization.ReadWrite.All", "Organization.Read.All",
    "Mailbox.Migration", "User.Read.All", "User.ReadBasic.All", "MailboxSettings.Read", "Mail.Send", "Calendars.Read", "Contacts.Read", "Mail.Read", "Mail.ReadWrite", 
    "Contacts.ReadWrite", "MailboxSettings.ReadWrite", "Tasks.Read", "Tasks.ReadWrite", "Calendars.ReadWrite.All", "Calendars.Read.All", "Place.Read.All")

$Script:SMTPPermissions = @("SMTP.Send", "SMTP.SendAsApp")

$Script:EWSPermissions = @("EWS.AccessAsUser.All", "full_access_as_app")

$Script:POPPermissions = @("POP.AccessAsUser.All", "POP.AccessAsApp")

$Script:IMAPPermissions = @("IMAP.AccessAsUser.All", "IMAP.AccessAsApp")

$Script:ReportingServicesPermissions = @("ReportingWebService.Read", "ReportingWebService.Read.All")

$Script:ExchangePowerShellPermissions = @("Exchange.Manage","Exchange.ManageAsApp")

#######################
# BEGIN MAIN FUNCTION #
#######################

Function Export-AppReg_APIPermissions
{
   
    Begin
    {
        Write-Host "-------------------------------------------"
        Write-Host "- SCRIPT STARTED AT: $($Script:StartTime)  -"
        Write-Host "-------------------------------------------"

        # Call function to Test all the Input parameters and set the required script variables
        Get-InputParameters

        # Call function to Test if the required Modules are installed (based on the script variables)
        Test-InstalledModules

        # Call function to create the output file and stream
        Create-OutputFile
    }

    Process
    {
        # connect to Azure AD and retrieve the App registrations
        Get-AppRegistrationsFromAzureAD
        

        $appCounter = 0

        # Loop through all the Application registrations
        Foreach ($AppReg in $Script:AzureAppRegs)
        {
            $Script:csvAppInfo = ""
            $Script:csvOutput = ""
            $appCounter++
            # Save the App registration Info for output to CSV file
            $Script:csvAppInfo = $Script:csvAppInfo + "" + $appCounter + "" + $Script:Tab + "" + $AppReg.AppId + "" + $Script:Tab + "" + $AppReg.DisplayName

            Write-Host "*********************************************"
            Write-Host "* $($appCounter)."
            Write-Host "* DisplayName: $($AppReg.DisplayName)"
            Write-Host "* AppID: $($AppReg.AppId)"
            Write-Host "*********************************************"
            Write-Host "Enumerating API Resource Access Groups"
            Write-Host ""
            $ReqResAccessCounter = 0

            # Catch the App Registrations that do not have any API Permisisons granted and save output
            # Using 'None' for the resource and permission details to keep the csv file format correctly
            If($($AppReg.RequiredResourceAccess) -eq $Null)
            {
                Save-BlankApplicationRegistration
            }
            else
            {
                Foreach ($ReqResAccess in $AppReg.RequiredResourceAccess)
                {

                    # Check if we are part of the Messaging API Resource Groups

                    $ReqResAccessCounter++
                    $AzureADServicePrincipal = $null

                    # Check to see if the Service Principal Object already exists in the Cache
                    if($Script:CacheServicePrincipal.ContainsKey($ReqResAccess.ResourceAppId))
                    {
                        $AzureADServicePrincipal = $Script:CacheServicePrincipal.Get_Item($ReqResAccess.ResourceAppId)
                    }else
                    {
                        # If the Service Principal object doesn't exist in the cache, retrieve it from Azure AD and add it to the cache
                        try
                        {
                            # Check if we are connected using MgGraph of using AzureAD module
                            # Depending, run either Get-MgServicePrincipal or Get-AzureADServicePrincipal cmdlet
                            switch ($UsePSModule) 
                            {
                                'Microsoft.Graph' 
                                    {
                                        $AzureADServicePrincipal = Get-MgServicePrincipal -All | Where-Object {$_.AppId -eq "$($ReqResAccess.ResourceAppId)"}
                                    }
                                'AzureAD' 
                                    {
                                        $AzureADServicePrincipal = Get-AzureADServicePrincipal -All $true | Where-Object {$_.AppId -eq "$($ReqResAccess.ResourceAppId)"}
                                    }
                              }

                            # Add the object to cache for faster processing
                            $Script:CacheServicePrincipal.Add($ReqResAccess.ResourceAppId, $AzureADServicePrincipal)
                        }catch [system.exception]
                            {
                                Write-Host "Error retrieving Service Principal for $($ReqResAccess.ResourceAppId), exiting script" -ForegroundColor Red                                
                                Exit
                            }
                    }
                    Write-Host "$($Script:Tab)$($ReqResAccessCounter) $($AzureADServicePrincipal.DisplayName) - $($ReqResAccess.ResourceAppId)"
                    Write-Host ""
                    $Script:csvResourceInfo = $AzureADServicePrincipal.DisplayName + "" + $Script:Tab + "" + $ReqResAccess.ResourceAppId
            
                    $Script:csvOutput = ""
                    Foreach ($ResourceAccess in $ReqResAccess.ResourceAccess)
                    {

                        $ProcessResourceAccessGroup = $flase
                        # Add a check to see if we are exporting the API permissions for the current ResourceAccess Group
                        # This is to avoid enumerating API Permissions in resource groups like PowerBi, etc...
                        # Except if we are exporting 'All" API Permissions
                        switch ($ExportAPIPermissions) 
                            {
                                'All' 
                                    {
                                        # We are exporting ALL API  Permissions so we will check every permission in every ResourceAccess Group
                                        $ProcessResourceAccessGroup = $true
                                        
                                    }
                                'OutlookRESTv2' 
                                    {
                                        # We are exporting OutlookRESTv2 permissions only, check if the current ResourceAccess Group 
                                        # is the 'Office 365 Exchange Online' Resource Access Group, if so process it
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            $ProcessResourceAccessGroup = $true
                                        }
                                    }
                                'Graph' 
                                    {
                                        # We are exporting Graph API permissions only, check if the current ResourceAccess Group 
                                        # is the 'Microsoft Graph' Resource Access Group, if so process it
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            $ProcessResourceAccessGroup = $true
                                        }
                                    }

                                'O365MgmtAPI' 
                                    {
                                        # We are exporting Graph API permissions only, check if the current ResourceAccess Group 
                                        # is the 'Office 365 Management APIs' Resource Access Group, if so process it
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:O365ManagementAPI)
                                        {
                                            $ProcessResourceAccessGroup = $true
                                        }
                                    }
                                default 
                                    {
                                        # This is the catch all were we are exporting POP, IMAP, SMTP, ExchangePowerShell, ReportingServices permissions only
                                        #check if the current ResourceAccess Group is either the 'Office 365 Exchange Online' or 'Microsoft Graph' Resource Access Group, if so process it
                                        
                                        if(([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph) -or ([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline))
                                        {
                                            $ProcessResourceAccessGroup = $true
                                        }
                                    }
                                }
                            #End of switch
                        
                        #Only process the ResourceAccess group when this has been set to True.
                        If($ProcessResourceAccessGroup)
                        {

                            ###########################################
                            # Scope are the Delegated API Permissions #
                            ###########################################
                            If($($ResourceAccess.Type) -eq 'Scope')
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                {
                            $AppScope = $null
                            # Check if we are conencted using MgGraph of using AzureAD module
                            # Depending, either use the Oauth2PermissionScopes or the Oauth2Permissions property
                            switch ($UsePSModule) 
                            {
                                'Microsoft.Graph' 
                                    {
                                        $AppScope = $AzureADServicePrincipal.Oauth2PermissionScopes | Where-Object {$_.Id -eq "$($ResourceAccess.Id)"}
                                    }
                                'AzureAD' 
                                    {
                                        $AppScope = $AzureADServicePrincipal.Oauth2Permissions | Where-Object {$_.Id -eq "$($ResourceAccess.Id)"}
                                    }
                              }
                            Write-Host "$($Script:Tab)$($Script:Tab)AdminConsentDisplayName : $($AppScope.AdminConsentDisplayName)"
                            Write-Host "$($Script:Tab)$($Script:Tab)Type : $($AppScope.Type)"
                            Write-Host "$($Script:Tab)$($Script:Tab)Value : $($AppScope.Value)"
                            Write-Host ""
                            # Create the Output string
                            $Script:csvPermissionInfo = $AppScope.AdminConsentDisplayName + "" + $Script:Tab +"" + $AppScope.Type + "" + $Script:Tab +"" + $AppScope.Value
                            $Script:csvOutput = $Script:csvAppInfo + "" + $Script:Tab + "" + $Script:csvResourceInfo + "" + $Script:Tab + "" + $Script:csvPermissionInfo
                            
                            # Switch to see what API permissions we are exporting
                            switch ($ExportAPIPermissions) 
                            {
                                'All' 
                                    {
                                        # We are exporting ALL API  Permissions
                                        Add-Content $Script:OutputStream $Script:csvOutput
                                    }
                                'OutlookRESTv2' 
                                    {
                                        # We are exporting OutlookRESTv2 permissions only
                                        # Check if the 'Delegated' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # See if the 'Delegated' Permission exists in the Array containing all Outlook REST API Application Permissions
                                            if($AppScope.Value -in $Script:Delegated_OutlookRESTPermissions)
                                            {
                                                # Save Output to file
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        }
                                    }
                                'EWS' 
                                    {
                                        # We are exporting EWS permissions only
                                        # Important: The EWS Delegated API permission can be granted through both resource groups 'Office 365 Exchange Online' or the 'Microsoft Graph'
                                        # Check if the 'Delegated' permission is part of the 'Office 365 Exchange Online' or the 'Microsoft Graph' resource group
                                        
                                        if(([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline) -or ([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph))
                                        {
                                            # Check if the 'Delegated' API Permissions is EWS Delegated EWS.AccessAsUser.All
                                            if($AppScope.Value -in $Script:EWSPermissions)
                                            {
                                                # Save Output to file
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        }
                                    }
                                'Graph'
                                    {
                                        # We are exporting Graph permissions only
                                        # Check if the 'Delegated' permission is part of the 'Microsoft Graph' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            # Save Output to file
                                            Add-Content $Script:OutputStream $Script:csvOutput
                                        }
                                    }
                                'O365MgmtAPI'
                                    {
                                        # We are exporting O365 Management API permissions only
                                        # Check if the 'Delegated' permission is part of the 'Office 365 Management APIs' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:O365ManagementAPI)
                                        {
                                            # Save Output to file
                                            Add-Content $Script:OutputStream $Script:csvOutput
                                        }
                                    }
                                'POP'
                                    {
                                        # We are exporting POP permissions only
                                        # Check if the 'Delegated' permission is part of the 'Microsoft Graph' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            # Check if the Application Scope exists in the POP API Permissions Array
                                            if($AppScope.Value -in $Script:POPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'IMAP'
                                    {
                                        # We are exporting IMAP permissions only
                                        # Check if the 'Application' permission is part of the 'Microsoft Graph' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            # Check if the Application Scope exists in the IMAP API Permissions Array
                                            if($AppScope.Value -in $Script:IMAPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'SMTP'
                                    {
                                        # We are exporting SMTP permissions only
                                        # Check if the 'Application' permission is part of the 'Microsoft Graph' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            # Check if the Application Scope exists in the SMTP API Permissions Array
                                            if($AppScope.Value -in $Script:SMTPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'ReportingServices'
                                    {
                                        # We are exporting Reporting Services API permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Scope exists in the ReportingServices API Permissions Array
                                            if($AppScope.Value -in $Script:ReportingServicesPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'ExchangePowerShell'
                                    {
                                        # We are exporting Exchange PowerShell permissions only
                                        # Check if the 'Delegated' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Scope exists in the Exchange PowerShell Permissions Array
                                            if($AppScope.Value -in $Script:ExchangePowerShellPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                #End of Switch
                            }
                        }
                            ############################################
                            # Role are the Application API Permissions #
                            ############################################

                            If ($($ResourceAccess.Type) -eq 'Role')
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    {
                            # Property AppRoles is the same in AzureAD and MgGraph so no need to check how we are connected
                            $AppRole = $AzureADServicePrincipal.AppRoles | Where-Object {$_.Id -eq "$($ResourceAccess.Id)"}

                            Write-Host "$($Script:Tab)$($Script:Tab)DisplayName : $($AppRole.DisplayName)"
                            Write-Host "$($Script:Tab)$($Script:Tab)Value : $($AppRole.Value)"
                            Write-Host "$($Script:Tab)$($Script:Tab)AllowedMemberTypes : $($AppRole.AllowedMemberTypes)"
                            Write-Host ""

                            $Script:csvPermissionInfo = $AppRole.DisplayName + "" + $Script:Tab + "" + $AppRole.AllowedMemberTypes + "" + $Script:Tab + "" + $AppRole.Value
                            $Script:csvOutput = $Script:csvAppInfo + "" + $Script:Tab + "" + $Script:csvResourceInfo + "" + $Script:Tab + "" + $Script:csvPermissionInfo

                            # Switch to see what permissions we are exporting
                            switch ($ExportAPIPermissions) 
                            {
                                'All' 
                                    {
                                        # We are exporting ALL API Permissions
                                        # Save Output to file
                                        Add-Content $Script:OutputStream $Script:csvOutput
                                    }
                                'OutlookRESTv2' 
                                    {
                                        # We are exporting Outlook REST permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # See if the Application Role exist in the Array containing all Outlook REST API Application Permissions
                                            if($AppRole.Value -in $Script:Application_OutlookRESTPermissions)
                                            {
                                                # Write-Host "$($AppRole.Value) exists in Array, exporting permission" -ForegroundColor Green
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        }
                                    }
                                'EWS' 
                                    {
                                        # We are exporting EWS permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the EWS API Permissions Array
                                            if($AppRole.Value -in $Script:EWSPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'Graph'
                                    {
                                        # We are exporting Graph permissions only
                                        # Check if the 'Application' permission is part of the 'Microsoft Graph' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:MicrosoftGraph)
                                        {
                                            # Save Output to file
                                            Add-Content $Script:OutputStream $Script:csvOutput
                                        }
                                    }
                                'O365MgmtAPI'
                                    {
                                        # We are exporting O365 Management API permissions only
                                        # Check if the Application Role is part of the 'Office 365 Management APIs' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:O365ManagementAPI)
                                        {
                                            # Save Output to file
                                            Add-Content $Script:OutputStream $Script:csvOutput
                                        }
                                    }
                                'POP'
                                    {
                                        # We are exporting POP permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the POP API Permissions Array
                                            if($AppRole.Value -in $Script:POPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'IMAP'
                                    {
                                        # We are exporting IMAP permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the IMAP API Permissions Array
                                            if($AppRole.Value -in $Script:IMAPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'SMTP'
                                    {
                                        # We are exporting SMTP permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the SMTP API Permissions Array
                                            if($AppRole.Value -in $Script:SMTPPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'ReportingServices'
                                    {
                                        # We are exporting Reporting Services API permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the ReportingServices API Permissions Array
                                            if($AppRole.Value -in $Script:ReportingServicesPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                'ExchangePowerShell'
                                    {
                                        # We are exporting Exchange PowerShell permissions only
                                        # Check if the 'Application' permission is part of the 'Office 365 Exchange Online' resource group
                                        if([string]$ReqResAccess.ResourceAppId -eq $Script:Office365ExchangeOnline)
                                        {
                                            # Check if the Application Role exists in the Exchange PowerShell Permissions Array
                                            if($AppRole.Value -in $Script:ExchangePowerShellPermissions)
                                            {
                                                Add-Content $Script:OutputStream $Script:csvOutput
                                            }
                                        
                                        }
                                    }
                                #End of Switch
                            }
                        }
                        }
                    }
                }
            }
        }

    }

    End
    {
        $EndTime = '{0:MM/dd/yyyy HH:mm:ss}' -f [DateTime]::Now
        Write-Host "-------------------------------------------"
        Write-Host "- SCRIPT FINISHED AT: $EndTime -"
        Write-Host "-------------------------------------------"
    }
}

########################################
# BEGIN DEFINITION OF HELPER FUNCTIONS #
########################################

# Helper function that validates all the Input Parameters and sets the Script wide variables
Function Get-InputParameters
{
    Write-Host "- Parameters:"
    Write-Host "-------------"
    # Use a switch to display which PowerShell Module will be used to connect to Azure AD
    # use default value to catch unexpected error
    switch ($UsePSModule) 
    {
        'Microsoft.Graph' 
            {
                Write-Host "- Script will use 'Microsoft.Graph' module to connect to Azure AD"
            }
        'AzureAD' 
            {
                Write-Host "- Script will use 'AzureAD' module to connect to Azure AD" 
            }
        default 
            { 
                Write-Host 'Unexpected Error: entering Default for UserPSModule switch, exiting script' -ForegroundColor Red 
                Exit 
            }
    }
    # Use a switch to display which API Permissions will be exported to the output file (csv)
    # use default value to catch unexpected error
    switch ($ExportAPIPermissions) 
    {
        'All' 
            {
                Write-Host "- Exporting All API Permissions"
            }
        'OutlookRESTv2' 
            {
                Write-Host "- Exporting Only Outlook REST API Permissions"
            }
        'EWS' 
            {
                Write-Host "- Exporting Only EWS API Permissions"
            }
        'Graph'
            {
                Write-Host "- Exporting Only Graph API Permissions"
            }
        'O365MgmtAPI'
            {
                Write-Host "- Exporting Only O365 Management API Permissions"
            }
        'POP'
            {
                Write-Host "- Exporting Only POP API Permissions"
            }
        'IMAP'
            {
                Write-Host "- Exporting Only IMAP API Permissions"
            }
        'SMTP'
            {
                Write-Host "- Exporting Only SMTP API Permissions"
            }
        'ReportingServices'
            {
                Write-Host "- Exporting Only ReportingServices API Permissions"
            }
        'ExchangePowerShell'
            {
                Write-Host "- Exporting Only Exchange PowerShell API Permissions"
            }
        default 
            { 
                Write-Host 'Unexpected Error: entering Default for ExportAPIPermissions switch, exiting script' -ForegroundColor Red 
                Exit 
            }
    }

    Write-Host "- Output Directory: $($OutputPath)"
    Write-Host "- Output File Name: $($Script:FileName)"
    Write-Host "-------------------------------------------"
}

# Helper function that validates if the required modules are installed
Function Test-InstalledModules
{
    # no longer needed to test for default value
    switch ($UsePSModule) 
    {
        'Microsoft.Graph' 
            {
                #Test if the required module is installed, if not exit the script and print a help message
                if(Get-InstalledModule -Name Microsoft.Graph -MinimumVersion 2.6.1) 
                {
                    Write-Host "- Module 'Microsoft.Graph' with Minimum version 2.6.1 is installed"
                } 
                else {
                    Write-Host "- This script requires 'Microsoft.Graph' module with Minimum version 2.6.1" -ForegroundColor Red
                    Write-Host "- Please install the required 'Microsoft.Graph' module from the PSGallery repository by running command:" -ForegroundColor Red
                    Write-Host "- install-module -Name Microsoft.Graph -MinimumVersion '2.6.1' -Repository:PSGallery" -ForegroundColor Red
                    Exit
                }
            }
        'AzureAD' 
            {
                # Test if user is using Windows PowerShell 7 as this the AzureAD module is not supported in PowerShell v7
                if($PSVersionTable.PSVersion.Major -eq 7)
                {
                    Write-Host "- This AzureAD module cannot be used in PowerShell 7" -ForegroundColor Red
                    Write-Host "- current PowerShell version: $($PSVersionTable.PSVersion.ToString())" -ForegroundColor Red
                    Write-Host "- Please run this script in a PowerShell version 5" -ForegroundColor Red
                    exit
                }
                # Test if the required module is installed, if not exit the script and print a help message
                if(Get-InstalledModule -Name AzureAD -MinimumVersion 2.0.2.182) 
                {
                    Write-Host "- Module 'AzureAD' with Minimum version 2.0.2.182 is installed"
                } 
                else {
                    Write-Host "- This script requires 'AzureAD' module with Minimum version 2.0.2.182" -ForegroundColor Red
                    Write-Host "- Please install the required 'AzureAD' module from the PSGallery repository by running command:" -ForegroundColor Red
                    Write-Host "- install-module -Name AzureAD -MinimumVersion '2.0.2.182' -Repository:PSGallery" -ForegroundColor Red
                    Exit
                }
            }
    }

}

# Helper function that creates the Output csv file and Output stream used to save the data
Function Create-OutputFile
{
    # Create the output file
    # If an Output path is defined via the parameter, first check if the provided Output Path exists, if not exit the script
    if(!(Test-Path -Path $OutputPath))
    {
        Write-Error "The provided OutputPath does not exist, exiting script" -ForegroundColor Red
        Exit
    }
    else
    {
        # The path exists, so creating the Output file
        $Script:OutputStream = New-Item -Path $OutputPath -Type file -Force -Name $($Script:FileName) -ErrorAction Stop -WarningAction Stop
        # Add the header to the csv file
        $strCSVHeader = "Index" + $Script:Tab + "AppID" + $Script:Tab + "AppDisplayName" + $Script:Tab + "ResourceAccessID_DisplayName" + $Script:Tab + "ResourceAccessID" + $Script:Tab + "APIPerm_DisplayName" + $Script:Tab + "APIPerm_Type" + $Script:Tab + "APIPerm_Value"
        Add-Content $Script:OutputStream $strCSVHeader
    }
}

# Helper function that connects to AzureAD and gets the Application Registration objects
# First check if we connecting to AzureAD using either the Microsoft.Graph or AzureAD module and
# connects to AzureAD either via Connect-MgGraph or Connect-AzureAD cmdlet
# Retrieve the App Registrations either via Connect-MgGraph or Get-MgApplication cmdlet
Function Get-AppRegistrationsFromAzureAD
{

    # no longer needed to test for default value
    switch ($UsePSModule) 
    {
        'Microsoft.Graph' 
            {
                Write-Host "- Connecting using Microsoft.Graph module"
                # Connect to Microsoft Graph
                try
                {
                    Connect-MgGraph -Scopes Application.Read.All -NoWelcome -ContextScope Process -WarningAction:SilentlyContinue
                }catch [system.exception]
                    {
                        Write-Host "Error connecting to Microsoft Graph, exiting script" -ForegroundColor Red
                        Exit
                    } 
                # Retrieve all the App Registrations
                Write-Host "- Retrieving all Azure AD Application Registrations"
                try
                {
                    $Script:AzureAppRegs  = Get-MgApplication -All
                }catch [system.exception]
                    {
                        Write-Host "Error retrieving the Application registrations from Azure AD, exiting script" -ForegroundColor Red
                        Exit
                    }
            }
        'AzureAD' 
            {
                Write-Host "- Connecting using AzureAD module"
                # Connect to AzureAD
                try
                {
                    Connect-AzureAD -WarningAction:SilentlyContinue | Out-Null
                }catch [system.exception]
                    {
                        Write-Host "Error connecting to AzureAD, exiting script" -ForegroundColor Red
                        Exit
                    }
                # Retrieve all the App Registrations
                Write-Host "- Retrieving all Azure AD Application Registrations"
                try
                {
                    $Script:AzureAppRegs  = Get-AzureADApplication -All $true -WarningAction:SilentlyContinue
                }catch [system.exception]
                    {
                        Write-Host "Error retrieving the Application registrations from Azure AD, exiting script" -ForegroundColor Red
                        Exit
                    } 
            }
    }

    Write-Host ""
    Write-Host "--------------"
    Write-Host "- Found $($Script:AzureAppRegs.Count) App Registrations"
    Write-Host "--------------"
}

# Helper function that saves the App Registrations that do not have any API Permisisons granted and save output
# Using 'None' for the resource and permission details to keep the csv file format correctly
Function Save-BlankApplicationRegistration
{
    Write-Host "Application Registration does not have any API Permisisons assigned"
    Write-Host ""
    # Save blank Resource Access Info for output to CSV file
    $Script:csvResourceInfo = "None" + $Script:Tab + "None" + $Script:Tab + "None" + $Script:Tab + "None" + $Script:Tab + "None"
    $Script:csvOutput = $Script:csvAppInfo  + $Script:Tab + $Script:csvResourceInfo
    # Only save the application registrations with no permissions to file if we are exporting ALL API Permisisons
    If($ExportAPIPermissions -eq 'All')
    {
        # Save Output to file
        Add-Content $Script:OutputStream $Script:csvOutput
    }
}


########################
# CALL THE MAIN SCRIPT #
########################
Export-AppReg_APIPermissions

