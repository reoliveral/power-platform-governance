
## Identify Apps Using Per-App Licensing in Power Platform.

### 📜 PowerShell Script: [Get-PerAppLicensedApps.ps1](Get-PerAppLicensedApps.ps1)

This script connects to the **Power Platform Administration API** to identify **canvas apps configured to use per-app licensing**—those that consume Power Apps per-app plan passes instead of requiring individual per-user licenses across environments in the tenant.

It is intended for **Power Platform administrators** and **Center of Excellence (CoE)** teams who need visibility into per-app license consumption and app-level license governance.

### 📄 What the Script Does

The script retrieves and correlates data from the Power Platform Administration layer:

1. **Power Platform Administration PowerShell**
   Uses the `Microsoft.PowerApps.Administration.PowerShell` module to retrieve environment and app metadata, including:

   - Environment display name and ID
   - App display name and internal app ID
   - App owner (UPN and Object ID)
   - App creation and last modified timestamps
   - App status and type

2. **License Designation Detection**
   Inspects each app's internal properties to determine its license classification. Checks multiple known property paths to handle API version differences across tenants:

   - `licenseDesignation`
   - `planClassification`
   - `licenseType`
   - `appLicenseType`
   - `appPlanClassification`
   - Boolean fallbacks: `premiumRequired`, `bypassConsent`

3. **Reporting and Export**
   After scanning all environments (or a targeted one), the script:

   - Displays a summary table of apps using per-app licensing
   - Optionally exports full results to a **CSV file** for audit or CoE tracking
   - Supports a `–ShowAllApps` mode to list all apps with their detected license designation
   - Includes a `–Diagnose` mode that dumps raw app properties per environment, helping administrators identify the correct property name when results are unexpectedly empty

**Usage examples:**
### Run against all environments
```powershell
.\Get-PerAppLicensedApps.ps1
```
### Scan to a single environment by GUID
```powershell
 .\Get-PerAppLicensedApps.ps1 -EnvironmentName "00000000-0000-0000-0000-000000000000"
```
### Output all apps with their detected license designation, not only per-app ones and full path to export results as a `.csv` file. 
```powershell
 .\Get-PerAppLicensedApps.ps1 -ShowAllApps -ExportPath "C:\Reports\AllAppsWithLicense.csv"
```

### Dump raw `Internal.properties` of the first app per environment to identify the correct license property name for your tenant 
```powershell
 .\Get-PerAppLicensedApps.ps1 -Diagnose
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
