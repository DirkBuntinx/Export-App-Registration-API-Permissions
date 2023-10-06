# PowerShell Export App Registration API Permisisons
 PowerShell script to Export App Registration API Permisisons from Azure AD
 
 SYNOPSIS
Script to export the API Permissions for each Application Registration in Azure AD.
IMPORTANT: This is a READ ONLY script; it will only read information for AzureAD and make no modifications.

DESCRIPTION
This script will export the API Permissions for each Application Registration in Azure AD.  

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

EXAMPLE 1
.\Export-AppReg_APIPermissions_v1.0.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:EWS -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'EWS' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 2
.\Export-AppReg_APIPermissions_v1.0.ps1 -UsePSModule:AzureAD -ExportAPIPermissions:EWS -OutputPath:'C:\temp'
The will cause PowerShell to use the 'AzureAD' module to retrieve all the Application registrations from AzureAD and will export ALL the 'EWS' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 3
.\Export-AppReg_APIPermissions_v1.0.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:OutlookRESTv2 -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'Graph' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 4
.\Export-AppReg_APIPermissions_v1.0.ps1 -UsePSModule:AzureAD -ExportAPIPermissions:OutlookRESTv2 -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL the 'Graph' API Permissions to a file 
called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 5
.\Export-AppReg_APIPermissions_v1.0.ps1 -UsePSModule:Microsoft.Graph -ExportAPIPermissions:All -OutputPath:'C:\temp'
The will cause PowerShell to use the 'Microsoft.Graph' module to retrieve all the Application registrations from AzureAD and will export ALL API Permissions for ALL Resource Access Group 
to a file called "Export-AppReg_APIPermissions_<timestamp>.csv" in the 'C:\temp' directory
IMPORTANT: Take extra care of using the "All" value for Switch -ExportAPIPermissions as this will process every API Permission for ALL Resource Access groups for every Application registratioon in AzureAD, depending on the
number of Application registrations in your tenant, this process might take a very long time.
Example, this will export API permissions for Resource Access Groups like PowerBI, Intune,Dynamics, all Azure related, etc

