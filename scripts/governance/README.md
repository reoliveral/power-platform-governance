# **Power Platform Governance – PowerShell Scripts**

## Identify Orphaned Dataverse for Teams Environments  
### 📜 Powershell Script: [DetectOrphanedTeamsEnvAzureResource.ps1](DetectOrphanedTeamsEnvAzureResource.ps1)
This script connects to **Azure**, **Microsoft Teams**, and the **Power Platform Administration API** to identify **orphaned Dataverse for Teams environments**—those that remain in the tenant but are no longer linked to an active Microsoft Team or whose environment owner has left the company.

It is intended for **Power Platform administrators** and **Center of Excellence (CoE)** teams who need visibility and cleanup insights for Teams‑based Power Platform assets.

### 📄 What the Script Does

The script correlates data from three sources:

1. **Azure Resource Graph**  
   Uses a **Kusto Query Language (KQL) query** to retrieve Power Platform environment metadata from Azure, including:
   - Environment ID
   - Environment type (Dataverse for Teams)

2. **Microsoft Teams PowerShell**  
   Retrieves the list of existing Teams in the tenant to determine:
   - Active Teams
   - Team IDs (Group IDs)
   - Deleted or missing Teams

3. **Power Platform Administration PowerShell**  
   Retrieves environment‑level metadata, including:
   - Environment display name
   - Environment type
   - Dataverse for Teams linkage
   - Creation and ownership context

By correlating these datasets, the script identifies **Dataverse for Teams environments in cases where the related Team has been removed or the Teams group no longer has an owner**.

---
## Identify all connections and specify the permissions associated with each connection and environment.
### 📜 Powershell Script: [Get-AllConnectionsAndPermissions.ps1](Get-AllConnectionsAndPermissions.ps1)
This script connects to **Power Platform Administration API** to identify **connections and permissions**.
If the principal type is tenant, the connection was shared with all users in the tenant.

**Usage examples:**
### Run against all environments
```
.\Get-AllConnectionsAndPermissions.ps1
```

### Export to CSV
```
.\Get-AllConnectionsAndPermissions.ps1 -OutputCsvPath "C:\reports\connections.csv"
```

### Scope to specific environments (by display name or ID)
```
.\Get-AllConnectionsAndPermissions.ps1 -EnvironmentFilter "54c030e3-ee4d-e188-b519-71baf285ac19"
```
---
## Identify all custom connectors and specify the permissions associated with each connection and environment.
### 📜 Powershell Script: [Get-AllCustomConnectorsAndSharedConnections.ps1](Get-AllCustomConnectorsAndSharedConnections.ps1)
This script connects to **Power Platform Administration API** to identify **custom connectors and permissions**.
If the principal type is tenant, the connection was shared with all users in the tenant.

**Usage examples:**
### Run against all environments
```
.\Get-AllCustomConnectorsAndSharedConnections.ps1
```

### Export to CSV
```
.\Get-AllCustomConnectorsAndSharedConnections.ps1 -OutputCsvPath "C:\reports\connections.csv"
```

### Scope to specific environments (by display name or ID)
```
.\Get-AllCustomConnectorsAndSharedConnections.ps1 -EnvironmentFilter "54c030e3-ee4d-e188-b519-71baf285ac19"
```

This is designed for **Power Platform administrators** and **Center of Excellence (CoE)** teams who require insight into connections and sharing details.

---
## Artifacts and the Access Permissions for Each Environment

### 📜 PowerShell Script: [Invoke-PPEnvironmentAccessAudit.ps1](./Invoke-PPEnvironmentAccessAudit.ps1)

This script connects to **Power Platform Administration**, **Microsoft Graph**, and optionally **Dataverse** to inventory the users, groups, service accounts, and tenant-wide shares that have access to resources in a single Power Platform environment.

It is intended for **Power Platform administrators**, **environment owners**, and **Center of Excellence (CoE)** teams who need to understand who has access to apps, flows, connections, custom connectors, Dataverse security roles, Copilot Studio agents, and environment-related resources before applying security-group based environment access controls.

### 📄 What the Script Does

The script audits one Power Platform environment at a time and produces CSV reports that can be reviewed with business owners before restricting environment access.

The script collects access information from the following sources:

1. **Power Platform Canvas Apps**  
   Retrieves canvas apps in the selected environment and identifies:
   - App owners
   - Role assignments
   - Users with access
   - Groups with access
   - Apps shared with the entire tenant

2. **Power Automate Cloud Flows**  
   Retrieves cloud flows in the selected environment and identifies:
   - Flow creators
   - Flow owners and co-owners
   - User-based ownership
   - Group-based ownership

3. **Power Platform Connections**  
   Retrieves connections in the environment and identifies:
   - Connection owners
   - Service accounts that own connections
   - User accounts whose connections may keep flows running

4. **Custom Connectors**  
   Retrieves custom connectors in the environment and identifies:
   - Connector owners
   - User accounts responsible for custom connector access

5. **Microsoft Graph Group Expansion**  
   When the `-ExpandGroups` switch is used, the script resolves discovered Entra ID groups and retrieves:
   - Transitive group members
   - Member display names
   - Member user principal names
   - Group-to-user access relationships

6. **Dataverse Security Roles**  
   When the `-IncludeDataverse` switch is used, the script queries Dataverse and identifies:
   - Users with at least one Dataverse security role
   - Owner teams and access teams with security roles
   - Entra ID group-backed teams with security roles
   - Team members who inherit Dataverse access through team membership

7. **Copilot Studio Agents**  
   When Dataverse auditing is enabled, the script also identifies:
   - Copilot Studio agent owners
   - Users and groups each agent is shared with
   - Shared principals retrieved from Dataverse access records

8. **Solution Cloud Flows**  
   When Dataverse auditing is enabled, the script retrieves solution-aware cloud flows from the Dataverse workflow table and identifies:
   - Solution cloud flow owners
   - User ownership records
   - Unresolved team or owner references

### 📁 Output Files

The script creates an output folder for the selected environment and generates the following CSV files:

1. **UserAccess_Detail.csv**  
   Contains every artifact-level access record discovered by the script.

2. **UserAccess_Summary.csv**  
   Contains one row per distinct principal with a summary of why that user, group, or service account has access.

3. **TenantWideShares.csv**  
   Lists apps or resources that are shared with the entire tenant.

4. **GroupsEncountered.csv**  
   Lists all Entra ID groups discovered during the audit.

5. **GroupMembers.csv**  
   Lists transitive user members of discovered groups when `-ExpandGroups` is used.

6. **UnresolvedPrincipals.csv**  
   Lists object IDs that could not be resolved through Microsoft Graph or Dataverse.

7. **SecurityGroupSeed_<environment>.csv**  
   Generates an Entra ID bulk-import ready list of user principal names that can be used to seed a security group for the environment.

**Usage examples:**
```powershell
.\Invoke-PPEnvironmentAccessAudit.ps1 `
    -EnvironmentId "5f04367c-0b4d-e9cb-8685-f3a9ecc5cf22" `
    -IncludeDataverse `
    -ExpandGroups
```

---
### 🔧 Prerequisites
PowerShell 7+ (recommended) or Windows PowerShell 5.1.  
Appropriate Power Platform permissions, depending on the task:  
- Power Platform Admin  
- Dynamics 365 Admin  
- Global Admin (where required)  

**Required PowerShell modules**  
Each script declares its dependencies at the top or in accompanying documentation.  

### ⚠️ Disclaimer
These resources are provided “as‑is”, with no warranties.

- Always review scripts before running them.  
- Validate outputs carefully.  
- Test in non‑production environments first.  

**You are responsible for ensuring usage complies with your organization’s policies and all applicable laws and regulations.**

### ✅ Supportability and SLA
This library is open-source and is not a Microsoft-provided resource. As a result, there is no SLA or direct support available from Microsoft or the author for this open-source component.

### 📜 License  
This project is licensed under the MIT License.
See the LICENSE file for details.
