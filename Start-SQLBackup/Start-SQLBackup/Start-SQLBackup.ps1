<#
.SYNOPSIS
    Starts a BACKUP DATABASE/LOG session on a specified server for all databases on a given SQL Server
.DESCRIPTION
    This script will attempt to automate the backup of databases on a SQL Server, while also managing the directory structure of where you backup to. The script can either backup all the databases, or some of the databases through the use of the EXCLUDEDATABASES parameter and/or SYSTEMOBJECTS parameter.
.PARAMETER servername
    The hostname or FQDN of the server you want the backup to run on. This is a required parameter.
.PARAMETER instancename
    The instance name of the SQL server you want the backup to run on, defaults to "DEFAULT" for non-named instances. This is optional parameter.
.PARAMETER databasename
    The name of the database you want to back up. This is an optioanl parameter. If not supplied, all non-system databases will be backed up.
.PARAMETER excludedatabases
    A comma-seperated list of databases you DO NOT want to be backed up. This is an optional parameter.
.PARAMETER backuptype
    Options: Full, Log, or Differential
    The type of backup to be performed. This is a required parameter.
.PARAMETER backupLocation
    The path where the database backup will be stored. Can be a local path or a UNC path. This is a required parameter.
.PARAMETER backupFileName
    The name of the backup file. This is an optional parameter, and it CANNOT be used unless the databasename parameter is provided. If not supplied, the backups will be named as follows: DatabaseName_BackupType_YYYYMMDD_HHMMSS.extension
.PARAMETER copyOnly
    Sets the COPY_ONLY flag during the backup operation, so differential chains do not become broken. This is an optional parameter.
.PARAMETER splitFiles
    Specifices if the backup should be "split" (spanned) across seperate files. This is an optional parameter. This "may" increase the speed of the backup operation. If used, a folder will be created in the supplied backuplocation parameter named "BackupCollection_YYYYMMDD_HHMMSS" and each file will be named as follows:  DatabaseName_BackupType_YYYYMMDD_HHMMSS_FileNumber.extension
.PARAMETER systemObjects
    If supplied, master, msdb, and model will be backed up as part of your backup operation. This is an optional paramater. TempDB and distribution (if it exists) will not be backed up.
.PARAMETER seperateDirectories
    This parameter will handle creating seperate directories per database and per backup type at the supplied backuplocation parameter. For example, if your database name is "somedatabaes" and your backup location is C:\Backups and you use this parameter, the following folder structure will be used: C:\Backups\SomeDatabase\Full or C:\Backups\SomeDatabase\Logs depending on the type of backup performed. This is an optional parameter.
    Developer note: If you plan on using scripted restores, this parameter will make automating your restores 1000x easier!
.PARAMETER certificateLocation
.PARAMETER privateKeyPassword
.PARAMETER privateKeyLocation
.PARAMETER sendKeyTo
.PARAMETER smtpServer
.PARAMETER ssisdb
.PARAMETER ssisdbKeyPassword
.EXAMPLE
    .
.OUTPUTS
    .
.NOTES
    .
.CHANGELOG
    .
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $servername,
    [Parameter(Mandatory=$false)] [string] $instanceName = "DEFAULT",
    [Parameter(Mandatory=$false)] [string] $databaseName,
    [Parameter(Mandatory=$false)] [string[]] $excludeDatabases,
    [Parameter(Mandatory=$true)] [validateset('Full','Differential','Log')] [string] $backupType, 
    [Parameter(Mandatory=$true)] [string] $backupLocation,
    [Parameter(Mandatory=$false)] [string] $backupFileName,
    [Parameter(Mandatory=$false)] [switch] $copyOnly,
    [Parameter(Mandatory=$false)] [int] $splitFiles = 1,
    [Parameter(Mandatory=$false)] [switch] $systemObjects,
    [Parameter(Mandatory=$false)] [switch] $seperateDirectories,
    [Parameter(Mandatory=$false)] [string] $certificateLocation,
    [Parameter(Mandatory=$false)] [string] $privateKeyPassword,
    [Parameter(Mandatory=$false)] [string] $privateKeyLocation,
    [Parameter(Mandatory=$false)] [string] $sendKeyTo,
    [Parameter(Mandatory=$false)] [string] $smtpServer,
    [Parameter(Mandatory=$false)] [switch] $ssisdb,    
    [Parameter(Mandatory=$false)] [string] $ssisdbKeyPassword
)

function Make-StrongPassword([int]$length, [int]$nonalphanumericchars)
{
    if ($length -lt $nonalphanumericchars) { $nonalphanumericchars = $length }
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
    return [System.Web.Security.Membership]::GeneratePassword($length,$nonalphanumericchars)
}

$whoami = ([adsisearcher]"(samaccountname=$env:USERNAME)").FindOne().Properties.mail
Write-Verbose "This script is running as $whoami"

if ($certificateLocation -and !$privateKeyLocation)
{
    Throw "If you're backing up a certificate, you should be backing up with an encrypted private key."
}

if ($certificateLocation -and !$privateKeyPassword)
{
    $privateKeyPassword = Make-StrongPassword 20 10
    Write-Warning "User did not provide an encryption password, using automatically generated password..."
}

$keyAction = $null
if ($sendKeyTo)
{
    if (Test-Path $sendKeyTo)
    {
        Write-Verbose "Private key passwords will be written to a file at the following location: $privateKeyLocation"
        $keyAction = "file"
    } else {
        try
        {
            $toAddress = New-Object Net.Mail.MailAddress($sendKeyTo)
            Write-Verbose "Private key passwords will be emailed to: $privateKeyLocation"
            $keyAction = "email"
        }
        catch
        {
            Throw "Unable to resolve email address or file location for private key passwords. Please check your 'privateKyLocation' parameter and try again."
        } 
    }  
    if ($keyAction -eq "email" -and !$smtpServer)
    {
        throw "You have requested private keys be emailed to $sendKeyTo but you didn't provide an smtp server to send the email. Please add -smtpserver to your command and try again."
    }
}

if ((Test-Path $backupLocation) -eq $false)
{
    Throw "Unable to resolve backup location. Stopping."
}

if (!$databaseName -and $backupFileName)
{
    Throw "You can't use the backup file name parameter without specifying a database."
}

$sqlPath = "SQLSERVER:\SQL\" + $servername + "\" + $instanceName

if ($systemObjects)
{
    $dbs = Get-ChildItem ($sqlPath + "\Databases") -Force | Where-Object {$_.Name -ne "tempdb" -and $_.Name -ne "distribution"}
} else {
    $dbs = Get-ChildItem ($sqlPath + "\Databases")
}

if (!$ssisdb)
{
    $dbs = $dbs | Where-Object {$_.Name -ne "SSISDB"} 
}

if ($excludeDatabases)
{
    foreach ($e in $excludeDatabases)
    {
        $dbs = $dbs | Where-Object {$_.Name -ne $e}
    }
}

if ($databaseName)
{
    $dbs = $dbs | Where-Object {$_.Name -eq $databaseName}
}

foreach ($d in $dbs)
{
    $d.Refresh()
}


if ($backupType -eq "Log")
{
    $dbs = $dbs | Where-Object {$_.RecoveryModel -eq "Full"}
}

if ($databaseName)
{
    $dbs = $dbs | Where-Object {$_.Name -eq $databaseName}
}

if ($dbs.Length -eq 0)
{
    throw "Nothing to backup!"
}

foreach ($d in $dbs)
{
    $timestamp = Get-Date -UFormat "%Y%m%d_%H%M%S"
    $currentDBName = $d.Name
    $fullBackupLocation = $backupLocation
    if ($seperateDirectories)
    {
        
        $fullBackupLocation = $backupLocation + "\" + $servername + "\" + ($currentDBName.trim()) + "\" + $backupType
        if ((Test-Path $fullBackupLocation) -eq $false) {New-Item -ItemType Directory -Force -Path $fullBackupLocation | Out-Null }
    }

    Write-Verbose "Creating $backupType backup of database $currentDBName..."

    if ($backupType -eq "Full")
    {
        if ($d.DatabaseEncryptionKey.EncryptionState -eq "Encrypted")
        {
            Write-Verbose "Database is encrypted!"
            if ($certificateLocation)
            {
                $certificateName = $d.DatabaseEncryptionKey.EncryptorName
                $certificateLocation = $certificateLocation
                $privateKeyLocation = $privateKeyLocation
                if ($seperateDirectories)
                {
                    $certificateLocation = $certificateLocation + "\" + $servername + "\" + $currentDBName + "\Certificates"
                    $privateKeyLocation = $privateKeyLocation + "\" + $servername + "\" + $currentDBName + "\Keys"
                }
                if ((Test-Path $certificateLocation) -eq $false) {New-Item -ItemType Directory -Force -Path $certificateLocation  | Out-Null }
                if ((Test-Path $privateKeyLocation) -eq $false) {New-Item -ItemType Directory -Force -Path $privateKeyLocation  | Out-Null }
                Write-Verbose "Backing up database certificate $certificateName to $certificateLocation"
                $srv = new-object ('Microsoft.SqlServer.Management.Smo.Server') $servername
                $srv.Databases["master"].ExecuteNonQuery("backup certificate " + $certificateName + " to file = '" + ($certificateLocation + "\" + $certificateName + "_" + $timestamp + ".cert") + "' WITH PRIVATE KEY (FILE = '" + ($privateKeyLocation + "\" + $certificateName + "_" + $currentDBName + "_" + $timestamp + ".key") + "',  ENCRYPTION BY PASSWORD = '" + $privateKeyPassword + "')")
                if ($keyAction -eq "email")
                {
                    send-mailmessage -to $sendKeyTo -body "Hello, this is an automated email from $servername/$instancename. The $currentDBName database backup is encrypted, and requires a certifcate with private key to restore. The certifcate was backed up to $certificateLocation and the private key was saved to $privateKeyLocation. The private key was encrypted with the following password: $privateKeyPassword" -From $whoami -SmtpServer $smtpServer -Subject "Backup of server certificate and key for encrypted database $currentDBName from $servername/$instancename"
                }
                if ($keyAction -eq "file")
                {
                    Write-Verbose "Encrypting private key with the following password: $privateKeyPassword"
                    "This is an automated message from $servername/$instancename. The $currentDBName database backup is encrypted, and requires a certifcate with private key to restore. The certifcate was backed up to $certificateLocation and the private key was saved to $privateKeyLocation. The private key was encrypted with the following password: $privateKeyPassword" | Out-File -FilePath ($sendKeyTo + "\" + $currentDBName + "_" + $timestamp + "_PrivateKey.txt")
                }
            } else {
                $certLastBackup = (Get-ChildItem -Path ($sqlPath + "\databases\master\certificates") | Where-Object {$_.Name -eq  $d.DatabaseEncryptionKey.EncryptorName}).LastBackupDate
                Write-Warning "This database is encrypted and you didn't specify a certifcate path!"
                Write-Warning "I strongly, STRONGLY suggest you take care of that..."
            }
        }
        if ($d.Name -eq "SSISDB")
        {
            Write-Verbose "Backing up Integration Services Catalog: SSISDB..."
            Write-Verbose "... but is it really an integration services catalog?"
            $ssisConnectionString = "Data Source=$servername;Integrated Security=SSPI;" 
            $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $sqlConnectionString
            try 
            {
                $ssis = New-Object "Microsoft.SqlServer.Management.IntegrationServices.IntegrationServices" $sqlConnection
                if ($ssis.catalogs[$d.Name].name -eq $d.name)
                {
                    Write-Verbose "Yup, looks like an Integration Services catalog to me..."
                    $ssisBackupLocation = $backupLocation
                    if ($seperateDirectories)
                    {
                        $ssisBackupLocation = $backupLocation + "\" + $servername + "\" + $currentDBName + "\IntegrationServices"
                    }
                    if ((Test-Path $ssisBackupLocation ) -eq $false) {New-Item -ItemType Directory -Force -Path $ssisBackupLocation  | Out-Null }
                    Write-Verbose "Scripting ##MS_SSISServerCleanupJobLogin## user to..."
                    (Get-ChildItem -Path ($sqlPath + "\Logins") | Where-Object {$_.Name -eq "##MS_SSISServerCleanupJobLogin##"}).Script() | Out-File -FilePath ($ssisBackupLocation + "\ssisobjects_" + $timestamp + ".sql") -append
                    Write-Verbose "Scripting dbo.sp_ssis_startup stored procedure to..."
                    (Get-ChildItem -Path ($sqlPath + "\Databases\Master\StoredProcedures") | Where-Object {$_.Name -eq "sp_ssis_startup" -and $_.Schema -eq "dbo"}).Script() | Out-File -FilePath ($ssisBackupLocation + "\ssisobjects_" + $timestamp + ".sql") -append
                    Write-Verbose "Backing up master database key to..."
                    $srv = new-object ('Microsoft.SqlServer.Management.Smo.Server') $servername
                    $srv.Databases[$d.Name].ExecuteNonQuery("backup master key to file = '" + ($ssisBackupLocation + "\databasemasterkey_" + $timestamp + ".key") + "' encryption by password = '" + $ssisdbKeyPassword + "'")
                    Write-Verbose "Saving password to ... because you'll probably forget it otherwise"
                    $ssisdbKeyPassword | Out-File -FilePath ($ssisBackupLocation + "\databasemasterkey_password_" + $timestamp + ".txt") -append
                } else {
                    Write-Warning "Could not find suitably-named Integration Services catalog to back up... skipping this step"
                }
                
            } catch {write-host "can't connect to ssis catalog"}
        }
        $fullbackup = New-Object ("Microsoft.SqlServer.Management.Smo.Backup")
        if ($backupFileName)
        {
            $fullbackup.Devices.AddDevice(($fullBackupLocation + "\" + $backupFileName), "File")
        } 
        else
        {
            if ($splitFiles -gt 1)
            {
                if ((Test-Path ($fullBackupLocation + "\BackupCollection_" + $timestamp)) -eq $false) {New-Item -ItemType Directory -Force -Path ($fullBackupLocation + "\BackupCollection_" + $timestamp) | Out-Null }
                For ($devices = 1; $devices -le $splitFiles; $devices++)
                {
                    $fullbackup.Devices.AddDevice(($fullBackupLocation + "\BackupCollection_" + $timestamp + "\"  + $currentDBName + "_" + $backupType + "_" + $timestamp + "_" + $devices + ".bak"), "File")
                }
            }
            else
            {
                $fullbackup.Devices.AddDevice(($fullBackupLocation + "\" + ($currentDBName.trim()) + "_" + $backupType + "_" + $timestamp + ".bak"), "File")            
            }
        }
        $parameters = @{
            Path = $sqlPath
            MediaDescription = ("Full Backup from " + $servername + " of " + $currentDBName + " on " + (Get-Date))
            BackupDevice = $fullbackup.Devices
            Database = $currentDBName
            BackupAction = "Database"
            CompressionOption = "On"
            FormatMedia = $true
            Initialize = $true
            SkipTapeHeader = $true
        }

        if ($copyOnly)
        {
            $parameters.Add("CopyOnly",$true);
        }
        Backup-SqlDatabase @parameters
    }
    if ($backupType -eq "Differential")
    {
        Backup-SqlDatabase -Path $sqlPath -MediaDescription ("Differential Backup from " + $servername + " of " + $currentDBName + " on " + (Get-Date)) -BackupFile ($fullBackupLocation + "\" + $currentDBName + "_" + $backupType + "_" + $timestamp + ".dif")  -Database $currentDBName -BackupAction Database -Incremental -CompressionOption On -FormatMedia -Initialize -SkipTapeHeader
    }
    if ($backupType -eq "Log")
    {
        Backup-SqlDatabase -Path $sqlPath -MediaDescription ("Transaction Log Backup from " + $servername + " of " + $currentDBName + " on " + (Get-Date)) -BackupFile ($fullBackupLocation + "\" + $currentDBName + "_" + $backupType + "_" + $timestamp + ".trn") -Database $currentDBName -BackupAction Log -CompressionOption On -FormatMedia -Initialize -SkipTapeHeader
    }
    Write-Verbose "Backup completed!"
}