$script:EsentDllPath = "$env:SYSTEMROOT\Microsoft.NET\assembly\GAC_MSIL\microsoft.isam.esent.interop\v4.0_10.0.0.0__31bf3856ad364e35\Microsoft.Isam.Esent.Interop.dll"

Function Get-ESEDatabase {
	<#
		.SYNOPSIS
			Enumerates a Extensible Storage Engine (ESE) database, providing all tables and data contained within those tables.

		.DESCRIPTION
			The Get-ESEDatabase cmdlet starts a new sessions with the given ESE database. If the database is in a dirty shutdown state, the cmdlet will run a repair or restore operation.

			It is recommended that an offline copy of the database is used for enumeration so that no data in an active database is lost. The database is opened with the recovery option set to false to stop errors with page size conflicts.

		.EXAMPLE
			Get-ESEDatabase -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")

			Gets an array of PSCustomObjects where each item in the array is a complete set of table data. The array represents the tables in the WebCacheV01.dat database. The processes dllhost and taskhostw are stopped to free the database from use. The log prefix is set to V01 to be used with the esentutl utility for repair operations.

		.PARAMETER Path
			The path to the ESE database.

		.PARAMETER LogPrefix
			The prefix of the logs files for the database to be used with esentutl repair operations.

		.PARAMETER FutureTimeLimit
			Two column data types, 14 and 15 are not specifically defined in the ESE documentation. Sometimes these are datetime objects, and sometimes they are Int64. In order to properly translate these data types, a future limit is set on converting them to a DateTime.

			Input is converted to a DateTime if it is between 1 Jan 1970 and a TimeSpan defined by this parameter added to the current date. This defaults to 100 years.

		.PARAMETER PageSize
			The page size to be used in reading the database. This information can be specified or defaults to being read from the database file. The value must be a multiple of 1024.

		.PARAMETER ProcessesToStop
			Specify any processes that will be stopped to free the database from exclusive locks, even for readonly operations.

		.PARAMETER Recovery
			Sets the Microsoft.Isam.Esent.Interop.JET_param.Recovery option in the JetSetSystemParameter object when opening the database. This defaults to false to prevent errors with the PageSize setting.

			This parameter is the master switch that controls crash recovery for an instance. If this parameter is set to "On" then ARIES style recovery will be used to bring all databases in the instance to a consistent state in the event of a process or machine crash. If this parameter is set to "Off" then all databases in the instance will be managed without the benefit of crash recovery. That is to say, that if the instance is not shut down cleanly using JetTerm prior to the process exiting or machine shutdown then the contents of all databases in that instance will be corrupted.

			https://msdn.microsoft.com/en-us/library/microsoft.isam.esent.interop.jet_param(v=exchg.10).aspx

		.PARAMETER CircularLogging
			This parameter configures how transaction log files are managed by the database engine. When circular logging is off, all transaction log files that are generated are retained on disk until they are no longer needed because a full backup of the database has been performed. When circular logging is on, only transaction log files that are younger than the current checkpoint are retained on disk. The benefit of this mode is that backups are not required to retire old transaction log files.

			This defaults to true.

			https://msdn.microsoft.com/en-us/library/microsoft.isam.esent.interop.jet_param(v=exchg.10).aspx

		.PARAMETER Credential
			An optional credential used to connect to the database.

		.PARAMETER Force
			Bypasses the check to confirm the operation since it may modify the database causing data loss.

		.INPUTS
			System.String

		.OUTPUTS
			System.Management.Automation.PSCustomObject[]

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 6/27/2017
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[ValidateScript({Test-Path -Path $_})]
		[System.String]$Path,

		[Parameter()]
		[System.String]$LogPrefix = [System.String]::Empty,

		[Parameter()]
		[System.TimeSpan]$FutureTimeLimit = [System.TimeSpan]::FromDays(36500), #100 years

		[Parameter()]
		[ValidateScript({($_ % 1024) -eq 0})]
		[System.Int32]$PageSize = -1,

		[Parameter()] 
		[System.String[]]$ProcessesToStop = @(),

		[Parameter()]
		[System.Boolean]$Recovery = $false,

		[Parameter()]
		[System.Boolean]$CircularLogging = $true,

		[Parameter()]
		[ValidateNotNull()]
		[System.Management.Automation.Credential()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

		[Parameter()]
		[switch]$Force
	)

	Begin {
	}

	Process {
		[System.Int32]$private:Result = 0

		if (-not $Force) 
		{
			$Title = "Confirm action."
			$Message = "This command may modify the database by running a repair if it is in a dirty shutdown state. You may lose data. It is recommended that you use an offline copy. Are you sure you want to continue?"

			$Yes = New-Object System.Management.Automation.Host.ChoiceDescription("&Yes","Executes database query.")
			$No = New-Object System.Management.Automation.Host.ChoiceDescription("&No", "Quits the cmdlet.")
			$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)

			$private:Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0) 
		}

		[PSCustomObject[]]$Tables = @()

		if ($private:Result -eq 0) 
		{
			Write-Verbose -Message "Initiating new session to $Path."
			$Force = $true

			$DBSession = New-ESEDatabaseSession -Path $Path -LogPrefix $LogPrefix -PageSize $PageSize -ProcessesToStop $ProcessesToStop -Recovery $Recovery -CircularLogging $CircularLogging -Credential $Credential -Force -ErrorAction Stop
			Write-Verbose -Message "Successfully initiated new session to $Path."

			$Session = $DBSession.Session
			$DatabaseId = $DBSession.DatabaseId
			$Instance = $DBSession.Instance

			try 
			{
				Write-Verbose -Message "Getting table names"
				[System.String[]]$TableNames = Get-ESEDatabaseTableNames -Session $Session -DatabaseId $DatabaseId

				Write-Verbose -Message "Iterating Tables"

				foreach ($TableName in $TableNames) 
				{
					Write-Verbose -Message "Processing table $TableName."

					try 
					{
						$Tables += Get-ESEDatabaseTableData -Session $Session -DatabaseId $DatabaseId -TableName $TableName -FutureTimeLimit $FutureTimeLimit
					}
					catch [Exception] 
					{
						Write-Warning -Message $_.Exception.Message
					}
				}
			}
			finally 
			{
				Write-Verbose -Message "Closing database connection as the final step."
				Close-ESEDatabase -Instance $Instance -Session $Session -DatabaseId $DatabaseId -Path $Path -ErrorAction SilentlyContinue
			}

			Write-Output -InputObject $Tables
		}
	}

	End {	
	}
}

Function New-ESEDatabaseSession {
	<#
		.SYNOPSIS
			Builds a new session with an ESE database and opens the database for ReadOnly operations.

		.DESCRIPTION
			The New-ESEDatabaseSession cmdlet starts a new sessions with the given ESE database. If the database is in a dirty shutdown state, the cmdlet will run a repair or restore operation.

			The cmdlet adds the ESENT library from $env:SYSTEMDRIVE\Microsoft.NET\assembly\GAC_MSIL\microsoft.isam.esent.interop\v4.0_10.0.0.0__31bf3856ad364e35\Microsoft.Isam.Esent.Interop.dll

			It is recommended that an offline copy of the database is used for enumeration so that no data in an active database is lost. The database is opened with the recovery option set to false to stop errors with page size conflicts.

			The database should be closed with the Close-ESEDatabase cmdlet after all operations are complete using this Session and Instance.

		.EXAMPLE
			New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")

			Gets a PSCustomObject with the database instance, database id, database object, and path information of the database. This information can be used with additional cmdlets to read data in the database.
	
			The processes dllhost and taskhostw are stopped to free the database from use. The log prefix is set to V01 to be used with the esentutl utility for repair operations.

		.PARAMETER Path
			The path to the ESE database.

		.PARAMETER LogPrefix
			The prefix of the logs files for the database to be used with esentutl repair operations.

		.PARAMETER PageSize
			The page size to be used in reading the database. This information can be specified or defaults to being read from the database file. The value must be a multiple of 1024.

		.PARAMETER ProcessesToStop
			Specify any processes that will be stopped to free the database from exclusive locks, even for readonly operations.

		.PARAMETER Recovery
			Sets the Microsoft.Isam.Esent.Interop.JET_param.Recovery option in the JetSetSystemParameter object when opening the database. This defaults to false to prevent errors with the PageSize setting.

			This parameter is the master switch that controls crash recovery for an instance. If this parameter is set to "On" then ARIES style recovery will be used to bring all databases in the instance to a consistent state in the event of a process or machine crash. If this parameter is set to "Off" then all databases in the instance will be managed without the benefit of crash recovery. That is to say, that if the instance is not shut down cleanly using JetTerm prior to the process exiting or machine shutdown then the contents of all databases in that instance will be corrupted.

			https://msdn.microsoft.com/en-us/library/microsoft.isam.esent.interop.jet_param(v=exchg.10).aspx

		.PARAMETER CircularLogging
			This parameter configures how transaction log files are managed by the database engine. When circular logging is off, all transaction log files that are generated are retained on disk until they are no longer needed because a full backup of the database has been performed. When circular logging is on, only transaction log files that are younger than the current checkpoint are retained on disk. The benefit of this mode is that backups are not required to retire old transaction log files.

			This defaults to true.

			https://msdn.microsoft.com/en-us/library/microsoft.isam.esent.interop.jet_param(v=exchg.10).aspx

		.PARAMETER Credential
			An optional credential used to connect to the database.

		.PARAMETER Force
			Bypasses the check to confirm the operation since it may modify the database causing data loss.

		.INPUTS
			System.String

		.OUTPUTS
			System.Management.Automation.PSCustomObject

			This object contains the following members:

			[Microsoft.Isam.Esent.Interop.JET_INSTANCE] - The database instance being opened
			[Microsoft.Isam.Esent.Interop.JET_SESID] - The session being used to access the database, the session could have multiple instances opened in it, but in this case it is just the one
			[Microsoft.Isam.Esent.Interop.JET_DBID] - The database ID of the database instance being opened
			[System.String] - The specified path to the database

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 6/27/2017
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[ValidateScript({Test-Path -Path $_})]
		[System.String]$Path,

		[Parameter()]
		[System.String]$LogPrefix = "V01",

		[Parameter()]
		[System.Int32]$PageSize = -1,

		[Parameter()] 
		[System.String[]]$ProcessesToStop = @(),

		[Parameter()]
		[System.Boolean]$Recovery = $false,

		[Parameter()]
		[System.Boolean]$CircularLogging = $true,

		[Parameter()]
		[ValidateNotNull()]
		[System.Management.Automation.Credential()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

		[Parameter()]
		[switch]$Force,

		[Parameter()]
		[ValidateScript({Test-Path -Path $_})]
		[string]$EsentDllPath = $script:EsentDllPath
	)

	Begin {
	}

	Process {
		if(-not $Force) 
		{
			$Title = "Confirm action."
			$Message = "This command may modify the database by running a repair if it is in a dirty shutdown state. You may lose data. It is recommended that you use an offline copy. Are you sure you want to continue?"

			$Yes = New-Object System.Management.Automation.Host.ChoiceDescription("&Yes","Executes database query.")
			$No = New-Object System.Management.Automation.Host.ChoiceDescription("&No", "Quits the cmdlet.")
			$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)

			$private:Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0) 
		} 
		else 
		{
			$private:Result = 0
		}

		if ($private:Result -eq 0) 
		{
			$Tables = @()
			$Connect = [System.String]::Empty
			[System.String]$Password = [System.String]::Empty
			[System.String]$UserName = [System.String]::Empty

			if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) 
			{
				$UserName = $Credential.UserName

				[System.IntPtr]$IntPtr = [System.IntPtr]::Zero

				try 
				{     
					$IntPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)     
					$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($IntPtr)   
				}   
				finally 
				{     
					if ($IntPtr -ne $null -and $IntPtr -ne [System.IntPtr]::Zero) 
					{       
						[System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($IntPtr)     
					}   
				}
			}

			if ($ProcessesToStop -ne $null)
			{
				foreach ($Process in $ProcessesToStop) 
				{
					if ((Get-Process -Name $Process -ErrorAction SilentlyContinue) -ne $null) 
					{
						Write-Verbose -Message "Stopping process $Process"
						Stop-Process -Name $Process -ErrorAction SilentlyContinue | Out-Null
						Write-Verbose -Message "Process stopped"
					}
					else 
					{
						Write-Verbose -Message "Process $Process does not exist."
					}
				}
			}

			[System.Int32]$FileType = -1
			[Microsoft.Isam.Esent.Interop.Api]::JetGetDatabaseFileInfo($Path, [ref]$FileType, [Microsoft.Isam.Esent.Interop.JET_DbInfo]::FileType)
			[Microsoft.Isam.Esent.Interop.JET_filetype]$DBType = [Microsoft.Isam.Esent.Interop.JET_filetype]($FileType)

			Write-Verbose -Message "File type $DBType."
		
			if ($DBType -eq [Microsoft.Isam.Esent.Interop.JET_filetype]::Database) 
			{
				if ($PageSize -eq -1 -or ($PageSize % 1024 -ne 0)) 
				{
					[Microsoft.Isam.Esent.Interop.Api]::JetGetDatabaseFileInfo($Path, [ref]$PageSize, [Microsoft.Isam.Esent.Interop.JET_DbInfo]::PageSize)
				}

				Write-Verbose -Message "Page size $PageSize."

				[Microsoft.Isam.Esent.Interop.JET_INSTANCE]$Instance = New-Object -TypeName Microsoft.Isam.Esent.Interop.JET_INSTANCE
				[Microsoft.Isam.Esent.Interop.JET_SESID]$Session = New-Object -TypeName Microsoft.Isam.Esent.Interop.JET_SESID

				$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetSetSystemParameter($Instance, [Microsoft.Isam.Esent.Interop.JET_SESID]::Nil, [Microsoft.Isam.Esent.Interop.JET_param]::DatabasePageSize, $PageSize, $null)
				$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetSetSystemParameter($Instance, [Microsoft.Isam.Esent.Interop.JET_SESID]::Nil, [Microsoft.Isam.Esent.Interop.JET_param]::Recovery, [int]$Recovery, $null)
				$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetSetSystemParameter($Instance, [Microsoft.Isam.Esent.Interop.JET_SESID]::Nil, [Microsoft.Isam.Esent.Interop.JET_param]::CircularLog, [int]$CircularLogging, $null)

				[Microsoft.Isam.Esent.Interop.Api]::JetCreateInstance2([ref]$Instance, "Instance", "Instance", [Microsoft.Isam.Esent.Interop.CreateInstanceGrbit]::None)
				$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetInit2([ref]$Instance, [Microsoft.Isam.Esent.Interop.InitGrbit]::None)
				[Microsoft.Isam.Esent.Interop.Api]::JetBeginSession($Instance, [ref]$Session, $UserName, $Password)

				[Microsoft.Isam.Esent.Interop.JET_DBID]$DatabaseId = New-Object -TypeName Microsoft.Isam.Esent.Interop.JET_DBID

				try 
				{
					try 
					{
						$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetAttachDatabase($Session, $Path, [Microsoft.Isam.Esent.Interop.AttachDatabaseGrbit]::ReadOnly)
						$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetOpenDatabase($Session, $Path, $Connect, [ref]$DatabaseId, [Microsoft.Isam.Esent.Interop.OpenDatabaseGrbit]::ReadOnly)
					}
					catch [Exception] {

						Write-Verbose -Message $_.Exception.Message
						Write-Verbose -Message "Running recovery on $Path with log prefix $LogPrefix."

						& "$env:SystemRoot\System32\esentutl.exe" "/r" "$LogPrefix"

						try 
						{
							$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetAttachDatabase($Session, $Path, [Microsoft.Isam.Esent.Interop.AttachDatabaseGrbit]::ReadOnly)
							$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetOpenDatabase($Session, $Path, $Connect, [ref]$DatabaseId, [Microsoft.Isam.Esent.Interop.OpenDatabaseGrbit]::ReadOnly)
						}
						catch [Exception] 
						{
							Write-Verbose -Message "Recovery failed, running repair on $Path."
							& "$env:SystemRoot\System32\esentutl.exe" "/p" "$Path" "/o"
							$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetAttachDatabase($Session, $Path, [Microsoft.Isam.Esent.Interop.AttachDatabaseGrbit]::ReadOnly)
							$Temp = [Microsoft.Isam.Esent.Interop.Api]::JetOpenDatabase($Session, $Path, $Connect, [ref]$DatabaseId, [Microsoft.Isam.Esent.Interop.OpenDatabaseGrbit]::ReadOnly)
						}
					}
				}
				catch [Exception] 
				{
					Write-Verbose -Message $_.Exception.Message
					Write-Verbose -Message "Shutting down database due to exception."

					try 
					{
						[Microsoft.Isam.Esent.Interop.Api]::JetDetachDatabase($Session, $Path)					
					}
					finally
					{
						[Microsoft.Isam.Esent.Interop.Api]::JetEndSession($Session, [Microsoft.Isam.Esent.Interop.EndSessionGrbit]::None)
						[Microsoft.Isam.Esent.Interop.Api]::JetTerm($Instance)
						Write-Verbose -Message "Completed shut down successfully."
						throw $_.Exception
					}
				}
			}
			else 
			{
				throw "The path must be to a database, the selected path was a $DBType."
			}

			Write-Output -InputObject ([PSCustomObject]@{Instance=$Instance;Session=$Session;DatabaseId=$DatabaseId;Path=$Path})
		}
	}

	End {		
	}
}

Function Get-ESEDatabaseTableNames {
	<#
		.SYNOPSIS
			Gets the table names from an ESE database.

		.DESCRIPTION
			The Get-ESEDatabaseTableNames cmdlet uses an existing session to a database and reads all of the table names.

		.EXAMPLE
			$Session = New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")
			Get-ESEDatabaseTableNames -Session $Session.Session -DatabaseId $Session.DatabaseId

			Gets an array of table names in the database.

		.PARAMETER Session
			The Microsoft.Isam.Esent.Interop.JET_SESID session object.

		.PARAMETER DatabaseId
			The Microsoft.Isam.Esent.Interop.JET_DBID database Id object.

		.INPUTS
			None

		.OUTPUTS
			System.String[]

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 4/25/2016
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_SESID]$Session,

		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_DBID]$DatabaseId
	)

	Begin {
	}

	Process {
		Write-Output -InputObject ([Microsoft.Isam.Esent.Interop.Api]::GetTableNames($Session, $DatabaseId))
	}

	End {		
	}
}

Function Get-ESEDatabaseTableColumns {
	<#
		.SYNOPSIS
			Gets the column information for a specific table.

		.DESCRIPTION
			The Get-ESEDatabaseTableColumns cmdlet uses an existing session to a database and reads the columns of a specified table.

		.EXAMPLE
			$Session = New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")
			Get-ESEDatabaseTableNames -Session $Session.Session -DatabaseId $Session.DatabaseId | ForEach-Object {
				Get-ESEDatabaseTableColumns -Session $Session.Session -DatabaseId $Session.DatabaseId -TableName $_
			}

			Gets a List of ColumnInfo for each table in the specified database.

		.PARAMETER Session
			The Microsoft.Isam.Esent.Interop.JET_SESID session object.

		.PARAMETER DatabaseId
			The Microsoft.Isam.Esent.Interop.JET_DBID database Id object.

		.PARAMETER TableName
			The name of the table to get the column information from.

		.INPUTS
			None

		.OUTPUTS
			 Microsoft.Isam.Esent.Interop.ColumnInfo[]

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 6/27/2017
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_SESID]$Session,

		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_DBID]$DatabaseId,

		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$TableName
	)

	Begin {
	}

	Process {
		[Microsoft.Isam.Esent.Interop.Table]$Table = New-Object -TypeName Microsoft.Isam.Esent.Interop.Table($Session, $DatabaseId, $TableName, [Microsoft.Isam.Esent.Interop.OpenTableGrbit]::None)
		Write-Output -InputObject ([Microsoft.Isam.Esent.Interop.ColumnInfo[]][Microsoft.Isam.Esent.Interop.Api]::GetTableColumns($Session, $Table.JetTableid))
	}

	End {		
	}
}

Function Get-ESEDatabaseTableData {
	<#
		.SYNOPSIS
			Gets all of the row information for a specified table.

		.DESCRIPTION
			The Get-ESEDatabaseTableData cmdlet uses an existing session to a database and reads all of the rows in a table.

		.EXAMPLE
			$Session = New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")
			Get-ESEDatabaseTableNames -Session $Session.Session -DatabaseId $Session.DatabaseId | ForEach-Object {
				Get-ESEDatabaseTableData -Session $Session.Session -DatabaseId $Session.DatabaseId -TableName $_
			}

			Gets all of the table data for each table in the database.

		.PARAMETER Session
			The Microsoft.Isam.Esent.Interop.JET_SESID session object.

		.PARAMETER DatabaseId
			The Microsoft.Isam.Esent.Interop.JET_DBID database Id object.

		.PARAMETER TableName
			The name of the table to get the column information from.

		.PARAMETER FutureTimeLimit
			Two column data types, 14 and 15 are not specifically defined in the ESE documentation. Sometimes these are datetime objects, and sometimes they are Int64. In order to properly translate these data types, a future limit is set on converting them to a DateTime.

			Input is converted to a DateTime if it is between 1 Jan 1970 and a TimeSpan defined by this parameter added to the current date. This defaults to 100 years.

		.INPUTS
			None

		.OUTPUTS
			System.Management.Automation.PSCustomObject

			The custom object contains the TableName, TableId, and an array of row data that are PSCustomObjects. The row data objects have properties corresponding to the table columns.

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 4/25/2016
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_SESID]$Session,

		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_DBID]$DatabaseId,

		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$TableName,

		[Parameter(Position = 3)]
		[System.TimeSpan]$FutureTimeLimit = [System.TimeSpan]::FromDays(36500)
	)

	Begin {
	}

	Process {
		Write-Verbose -Message "Getting table data for $TableName."
		
		try 
		{
			[Microsoft.Isam.Esent.Interop.Table]$Table = New-Object -TypeName Microsoft.Isam.Esent.Interop.Table($Session, $DatabaseId, $TableName, [Microsoft.Isam.Esent.Interop.OpenTableGrbit]::None)   

			$NewTable = @{Name=$Table.Name;Id=$Table.JetTableid;Rows=@()}
            
			[Microsoft.Isam.Esent.Interop.ColumnInfo[]]$Columns = [Microsoft.Isam.Esent.Interop.Api]::GetTableColumns($Session, $Table.JetTableid)

			if ([Microsoft.Isam.Esent.Interop.Api]::TryMoveFirst($Session, $Table.JetTableid)) 
			{
				do 
				{
					$NewTable.Rows += Get-ESEDatabaseTableRowData -Session $Session -TableId $Table.JetTableid -Columns $Columns -FutureTimeLimit $FutureTimeLimit
				} while ([Microsoft.Isam.Esent.Interop.Api]::TryMoveNext($Session, $Table.JetTableid))               
			}

			Write-Output -InputObject ([PSCustomObject]$NewTable)
		}
		catch [Exception] 
		{ 
			Write-Warning -Message $_.Exception.Message
		}
	}

	End {		
	}
}

Function Get-ESEDatabaseTableRowData {
	<#
		.SYNOPSIS
			Gets all the current row information.

		.DESCRIPTION
			The Get-ESEDatabaseTableRowData cmdlet uses an existing session to a database and reads the current row information.

			This cmdlet should be used in combination with the [Microsoft.Isam.Esent.Interop.Api]::TryMoveNext($Session, $Table.JetTableid)) command to iterate over the rows in the table.

		.EXAMPLE
			if ([Microsoft.Isam.Esent.Interop.Api]::TryMoveFirst($Session, $Table.JetTableid)) {
				do {
					Get-ESEDatabaseTableRowData -Session $Session -TableId $Table.JetTableid -Columns $Columns -FutureTimeLimit $FutureTimeLimit
				} while ([Microsoft.Isam.Esent.Interop.Api]::TryMoveNext($Session, $Table.JetTableid))               
			}

			Gets all of the table data for given table by iterating over each row and retrieving that data.

		.PARAMETER Session
			The Microsoft.Isam.Esent.Interop.JET_SESID session object.

		.PARAMETER TableId
			The Microsoft.Isam.Esent.Interop.JET_TABLEID table Id object.

		.PARAMETER Columns
			The set of columns to get information for in the row as System.Collections.Generic.List[Microsoft.Isam.Esent.Interop.ColumnInfo].

			If this input is $null or the default, the cmdlet enumerates the column information and uses all columns.

		.PARAMETER FutureTimeLimit
			Two column data types, 14 and 15 are not specifically defined in the ESE documentation. Sometimes these are datetime objects, and sometimes they are Int64. In order to properly translate these data types, a future limit is set on converting them to a DateTime.

			Input is converted to a DateTime if it is between 1 Jan 1970 and a TimeSpan defined by this parameter added to the current date. This defaults to 100 years.

		.INPUTS
			None

		.OUTPUTS
			System.Management.Automation.PSCustomObject

			The custom object contains a property and value for each column defined in the table and represents one row of data.

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 6/27/2017
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Position = 0, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_SESID]$Session,

		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_TABLEID]$TableId,

		[Parameter(Position = 2, Mandatory = $true)]
		[Microsoft.Isam.Esent.Interop.ColumnInfo[]]$Columns = $null,

		[Parameter(Position = 3)]
		[System.TimeSpan]$FutureTimeLimit = [System.TimeSpan]::FromDays(36500)
	)

	Begin {
	}

	Process {
		$Row = @{}

		if ($Columns -eq $null -or $Columns.Length -eq 0) 
		{
			$Columns = [Microsoft.Isam.Esent.Interop.Api]::GetTableColumns($Session, $Table.JetTableid)
		}

		foreach ($Column in $Columns) 
		{ 
			switch ($Column.Coltyp) 
			{
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Bit) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsBoolean($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::DateTime) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsDateTime($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::IEEEDouble) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsDouble($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::IEEESingle) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsFloat($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Long) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsInt32($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Binary) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsString($Session, $Table.JetTableid, $Column.Columnid, [System.Text.Encoding]::UTF8)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::LongBinary) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsString($Session, $Table.JetTableid, $Column.Columnid, [System.Text.Encoding]::UTF8)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::LongText) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsString($Session, $Table.JetTableid, $Column.Columnid, [System.Text.Encoding]::UTF8)
                            
					#Replace null characters which are 0x0000 unicode                                                     
					if (![System.String]::IsNullOrEmpty($Buffer)) {
						$Buffer = $Buffer.Replace("`0", "")
					}
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Text) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsString($Session, $Table.JetTableid, $Column.Columnid, [System.Text.Encoding]::UTF8)
                                
					#Replace null characters which are 0x0000 unicode                                                     
					if (![System.String]::IsNullOrEmpty($Buffer)) {
						$Buffer = $Buffer.Replace("`0", "")
					}
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Currency) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsString($Session, $Table.JetTableid, $Column.Columnid, [System.Text.Encoding]::UTF8)
                              
					#Replace null characters which are 0x0000 unicode                                                     
					if (![System.String]::IsNullOrEmpty($Buffer)) {
						$Buffer = $Buffer.Replace("`0", "")
					}
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::Short) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsInt16($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				([Microsoft.Isam.Esent.Interop.JET_coltyp]::UnsignedByte) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsByte($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				(14) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsInt32($Session, $Table.JetTableid, $Column.Columnid)
					break
				}
				(15) {
					$Buffer = [Microsoft.Isam.Esent.Interop.Api]::RetrieveColumnAsInt64($Session, $Table.JetTableid, $Column.Columnid)
							
					try {
						$DateTime = [System.DateTime]::FromBinary($Buffer)
						$DateTime = $DateTime.AddYears(1600)
                               
						if ($DateTime -gt (Get-Date -Year 1970 -Month 1 -Day 1) -and $DateTime -lt ([System.DateTime]::UtcNow.Add($FutureTimeLimit))) {
							$Buffer = $DateTime
						}
					}
					catch {}
							
					break							
				}
				default {
					Write-Warning -Message "Did not match column type to $_"
					$Buffer = [System.String]::Empty
					break
				}
			}

			$Row.Add($Column.Name, $Buffer)                               
		}

		Write-Output -InputObject ([PSCustomObject]$Row)
	}

	End {	
	}
}

Function Close-ESEDatabase {
	<#
		.SYNOPSIS
			Closes an open ESE database session.

		.DESCRIPTION
			The Close-ESEDatabase cmdlet closes and detaches the database. Then it closes the session and terminates the JET instance with the open session.

		.EXAMPLE
			$Session = New-ESEDatabaseSession -Path C:\Users\Administrator\AppData\Local\Microsoft\Windows\WebCache\WebCacheV01.dat -LogPrefix "V01" -ProcessesToStop @("dllhost","taskhostw")
			Close-ESEDatabase -Instance $Session.Instance -Session $Session.Session -DatabaseId $Session.DatabaseId -Path $Session.Path

			Closes the open database.

		.PARAMETER Instance
			The Microsoft.Isam.Esent.Interop.JET_INSTANCE instance object.

		.PARAMETER Session
			The Microsoft.Isam.Esent.Interop.JET_SESID session object.

		.PARAMETER DatabaseId
			The Microsoft.Isam.Esent.Interop.JET_DBID database Id object.

		.PARAMETER Path
			The path to the database file.

		.INPUTS
			None

		.OUTPUTS
			None

		.NOTES
			AUTHOR: Michael Haken
			LAST UPDATE: 6/27/2017
	#>
	[CmdletBinding()] 
	Param(
		[Parameter(Position = 0, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_INSTANCE]$Instance,

		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_SESID]$Session,

		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateNotNull()]
		[Microsoft.Isam.Esent.Interop.JET_DBID]$DatabaseId,

		[Parameter(Position = 3,Mandatory  =$true)]
		[ValidateScript({Test-Path -Path $_})]
		[System.String]$Path
	)

	Begin {
	}

	Process {
		Write-Verbose -Message "Shutting down database $Path due to normal close operation."
		[Microsoft.Isam.Esent.Interop.Api]::JetCloseDatabase($Session, $DatabaseId, [Microsoft.Isam.Esent.Interop.CloseDatabaseGrbit]::None)
		[Microsoft.Isam.Esent.Interop.Api]::JetDetachDatabase($Session, $Path)
		[Microsoft.Isam.Esent.Interop.Api]::JetEndSession($Session, [Microsoft.Isam.Esent.Interop.EndSessionGrbit]::None)
		[Microsoft.Isam.Esent.Interop.Api]::JetTerm($Instance)
		Write-Verbose -Message "Completed shut down successfully."
	}

	End {
	}
}