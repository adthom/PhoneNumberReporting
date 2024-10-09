# PhoneNumberReporting
This repository contains Azure Automation runbooks to synchronize Microsoft Teams phone numbers with a SharePoint list. The SharePoint list is configured with RBAC, and these runbooks maintain that state. This project does not modify Microsoft Teams and is read-only outside of the referenced SharePoint lists.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Deployment](#deployment)
- [Post Deployment](#postdeployment)
- [Runbooks](#runbooks)
    - [CreateLists.ps1](#createlistsps1)
    - [UpdateRBAC.ps1](#updaterbacps1)
    - [UpdateDIDReportAssignments.ps1](#updatedidreportassignmentsps1)
- [Helper Scripts](#helperscripts)
    - [CreatePnPAppRegistration.ps1](#createpnpappregistrationps1)
    - [AssignPermissions.ps1](#assignpermissionsps1)
    - [PopulateData.ps1](#populatedataps1)
- [Configuration](#configuration)
- [License](#license)

## Overview

These runbooks streamline the management of phone number data and permissions within an Azure environment. They use Microsoft Graph, Microsoft Teams, and SharePoint PnP PowerShell modules, leveraging Managed Identity for automation account access to both Teams and SharePoint.

## Prerequisites

- A SharePoint Site to host the required lists
- Site Owner permissions for the SharePoint site
- Microsoft Graph API permissions:
  - Application.ReadWrite.All
  - RoleManagement.ReadWrite.Directory
  - Sites.FullControl.All
- Azure Subscription and Resource Group with deployment permissions

## Deployment

To deploy this solution, follow these steps:

1. Deploy the solution using your preferred method. A [Bicep template](deploy.bicep) and [ARM template](template.json) are provided. Update the configuration via the template parameters.
2. Run the [CreatePnPAppRegistration](CreatePnPAppRegistration.ps1) script if you do not already have an app registration for SharePoint PnP PowerShell.
3. Run the [AssignPermissions](AssignPermissions.ps1) script to configure the needed permissions for the Managed Identity for the Azure Automation Account.
4. Run the [CreateLists](CreateLists.ps1) runbook from the automation account to configure the needed lists in the SharePoint site.
5. Optionally, run the [PopulateData](PopulateData.ps1) script to pre-populate the source list of all managed phone numbers. If not, the DID <-> Department Map list will need to be populated manually.

## Post Deployment

Configure any numbers to be monitored in the DID <-> Department Map List in the SharePoint site. Permissions are granted at the Department level based on the Report RBAC list values. The report view to be shared is the Phone Number Report list.

## Runbooks

### CreateLists.ps1

Creates the necessary SharePoint lists and fields for storing and managing phone number data.

### UpdateRBAC.ps1

Updates the role-based access control settings for phone number data, ensuring correct permissions based on the latest data. The Phone Number Report permissions are set at a row level, with Site Owners given "Full Control" and users/groups in the Report RBAC List granted Read permissions.

### UpdateDIDReportAssignments.ps1

Updates the report data with the current assignment status for each phone number in Microsoft Teams.

## Helper Scripts

### CreatePnPAppRegistration.ps1

Registers the necessary Azure AD applications and grants permissions to interact with SharePoint and Microsoft Graph.

### AssignPermissions.ps1

Assigns the necessary permissions to service principals and automation accounts to manage phone number data, ensuring correct roles for the automation account.

### PopulateData.ps1

Populates initial data into the relevant SharePoint lists, including the RBAC list and the DID to Department map, ensuring correct data mapping.

## Configuration

The following variables are exposed in the Automation Account to configure your deployment
The following variables are exposed in the Automation Account to configure your deployment:

- `DIDDepartmentMapName`: The name of the DID <-> Department Map SharePoint List (default: `DIDDepartmentMap`)
- `RBACListName`: The name of the RBAC SharePoint List (default: `ReportRBAC`)
- `ReportName`: The name of the Report SharePoint List (default: `PhoneNumberReport`)
- `SharePointDomain`: The subdomain of your SharePoint Online domain to use for the SharePoint Lists (default: `None`)
    - Example: For `contoso.sharepoint.com`, enter `contoso`
- `Site`: The site ID of the SharePoint Site to use for the SharePoint Lists (default: `None`)
    - Example: For `https://contoso.sharepoint.com/sites/PhoneNumberReport`, enter `PhoneNumberReport`
- `RBACLastRun`: The last time the RBAC runbook was run (default: `None`)
    - Do not set this manually
- `Environment`: The environment of the SharePoint tenant (default: `Production`)

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
