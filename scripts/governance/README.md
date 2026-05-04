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

## Identify all connections and specify the permissions associated with each connection and environment.
### 📜 Powershell Script: [Get-AllConnectionsAndPermissions.ps1](Get-AllConnectionsAndPermissions.ps1)
This script connects to **Power Platform Administration API** to identify **connections and permissions**.
If the principal type is tenant, the connection was shared with all users in the tenant.

**Usage examples:**
### Run against all environments
.\Get-AllConnectionsAndPermissions.ps1

### Export to CSV
.\Get-AllConnectionsAndPermissions.ps1 -OutputCsvPath "C:\reports\connections.csv"

### Scope to specific environments (by display name or ID)
.\Get-AllConnectionsAndPermissions.ps1 -EnvironmentFilter "54c030e3-ee4d-e188-b519-71baf285ac19"

This is designed for **Power Platform administrators** and **Center of Excellence (CoE)** teams who require insight into connections and sharing details.

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
