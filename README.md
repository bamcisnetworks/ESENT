# Extensible Storate Engine (ESE) Cmdlets

Provides PowerShell cmdlets to use with the built-in Extensible Storage Engine (ESE) aka JET Blue. The module utilizes the ManagedESENT .NET library and provides read-only access to existing ESENT databases.

## Usage

### 1
$DB = Get-ESEDatabase -Path "$env:USERPROFILE\AppData\Local\Temp\WebCache\WebCacheV01.dat" -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw") -Recovery $false -CircularLogging $true -Force

This call gets the contents of the IE web history for the current user. The session is automatically closed by the cmdlet.

### 2
$Session = New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")
Get-ESEDatabaseTableNames -Session $Session.Session -DatabaseId $Session.DatabaseId
Close-ESEDatabase -Instance $Session.Instance -Session $Session.Session -DatabaseId $Session.DatabaseId -Path $Session.Path

This call opens a new session, enumerates the table names in the database, and then closes the session.

## Revision History

### 1.0.0.0
Initial Release