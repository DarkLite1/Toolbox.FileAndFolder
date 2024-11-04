Function Add-QuotaWarningHC {
    <#
    .SYNOPSIS
        Add a quota warning on a folder which sends out an e-mail when the
        quota threshold is reached.

    .PARAMETER MailTo
        Valid options:
        - [Source Io Owner Email] : the user that is surpassing the threshold:
        - [Admin Email]           : the system administrator
        - bob@gmail.com           : a fixed e-mail address
    #>
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [ValidateRange(1, 100)]
        [Int]$Threshold,
        [Parameter(Mandatory)]
        [String]$MailTo,
        [Parameter(Mandatory)]
        [String]$MailSubject,
        [Parameter(Mandatory)]
        [String]$MessageText
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager
        $Quota = $FQTM.GetQuota($Path)

        Write-Verbose "'$Path' Add quota warning with threshold '$Threshold'"
        $Quota.AddThreshold($Threshold)
        $Action = $Quota.CreateThresholdAction($Threshold, 2)
        $Action.MailTo = $MailTo
        $Action.MailSubject = $MailSubject
        $Action.MessageText = $MessageText

        <#
            1 Execute a command or script.
            2 Send an email message.
            3 Log an event to the Application event log.
            4 Generate a report.
        #>

        $Quota.Commit()
    }
    Catch {
        throw "Failed adding the quota warning on '$Path' with threshold '$Threshold' on '$env:COMPUTERNAME': $_"
    }
}

Function Convert-StringToQuotaFlagHC {
    Param (
        $Name
    )

    Try {
        switch ($Name) {
            $null { $null; break }
            'Soft' { 0; break }
            'Hard' { 256; break }
            'Disabled (Soft)' { 512; break }
            'Disabled (Hard)' { 768; break }
            Default { "Unknown string '$_'" }
        }
    }
    Catch {
        throw "Failed converting quota limit type '$Number'"
    }
}

Function Convert-QuotaStringToSoftLimitHC {
    Param (
        $Name
    )

    Try {
        switch ($Name) {
            $null { $null; break }
            'Soft' { $true; break }
            'Hard' { $false; break }
            'Disabled (Soft)' { $true; break }
            'Disabled (Hard)' { $true; break }
            Default { "Unknown string '$_'" }
        }
    }
    Catch {
        throw "Failed converting quota string to softlimit for '$Number': $_"
    }
}

Function Copy-FilesHC {
    <#
        .SYNOPSIS
            Copy files from the source folder to the destination folder.

        .DESCRIPTION
            Copy files from the source folder to the destination folder. Duplicate file names in the destination folder will be overwritten. When the option 'Structure' is select, only the files within the root folder are copied, not the subfolders.

        .PARAMETER Source
            Can be a path or a file name. If a path is provided the complete content will be copied. When a file name is provided only that file will be copied.

        .PARAMETER Destination
            The destination path where the content of the source folder will be written. Existing files with the same name will be overwritten.

        .PARAMETER Structure
            When a 'Structure' is used when copying one file, the file is analyzed for it's creation date and the destination folder will change to that name 'Destination\CREATION DATE\File.txt'. If a folder path is used as 'Source' and the 'Structure' switch is used we only copy the files on the root level each to their own folder based on the creation date.
            Valid options are:
            - 'yyyy-MM-dd'

        .EXAMPLE
            Copy-FilesHC -Source 'E:\Fruits\Bananas.txt' -Destination 'E:\Basket'
            Copies the file 'Bananas.txt' to the folder 'E:\Basket\Bananas.txt'.

        .EXAMPLE
            Copy-FilesHC -Source 'E:\Share\Appels.txt' -Destination 'E:\Test2' -Structure 'yyyy-MM-dd'
            Copies the file 'Appels.txt' to the folder 'E:\Test2\2014-06-15\Appels.txt'.

        .EXAMPLE
            Copy-FilesHC -Source '\\Domain\Share' -Destination '\\Domain\Backup share' -Structure 'yyyy-MM-dd'
            Copies the all the files in the folder 'Share' to the folder '\\Domain\Backup share\2014-06-15\File', or any other date as it's based on the file 'creation date'.

        .EXAMPLE
            Copy-FilesHC -Source 'E:\Mails' -Destination 'E:\Archive'
            Copies everything from the folder 'Mails' to the folder 'Archive'.
         #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript( {
                if (Test-Path -LiteralPath $_) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [String]$Source,
        [Parameter(Mandatory, Position = 1)]
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [String]$Destination,
        [ValidateSet('yyyy-MM-dd')]
        [String]$Structure
    )

    Begin {
        Function Set-DestinationHC {
            <#
        .SYNOPSIS
            Creates the destination folder with the correct name.

        .DESCRIPTION
            Creates the destination folder with the correct name. The name is based on the file's creation date.

        .PARAMETER Structure
            The date format used for creating the folder name.

        .PARAMETER Folder
            The default destination folder.

        .PARAMETER File
            The file where we take the creation date from to generate the folder name.
        #>

            Param (
                $Structure,
                $Folder,
                $File
            )
            Process {
                if ($Structure) {
                    $Date = $(Get-ChildItem -LiteralPath $File).CreationTime.Date.ToString($Structure)

                    $Destination = Join-Path -Path $Folder -ChildPath $Date
                    if (!(Test-Path -PathType Container $Destination)) {
                        $null = New-Item -ItemType Directory $Destination
                    }
                    Write-Output $Destination
                }
                else {
                    Write-Output $Folder
                }
            }
        }
    }

    Process {
        if (Test-Path -LiteralPath $Source -PathType Container) {
            if ($Structure) {
                $Scr = Get-ChildItem -LiteralPath $Source -File
            }
            else {
                $Scr = Get-ChildItem -LiteralPath $Source -Recurse
            }
        }
        else {
            $Scr = Get-ChildItem -LiteralPath $Source
        }

        foreach ($S in $Scr.FullName) {

            if ($Structure) {
                $Target = Set-DestinationHC -Structure $Structure -Folder $Destination -File $S
            }
            else {
                $Target = $Destination
            }

            Copy-Item -LiteralPath $S -Destination $Target -Recurse -Force -PassThru |
            ForEach-Object {
                Write-Output "Copy-FilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | Copied: From: '$Source' To: '$($_.FullName)'"
                $a = $True
            }

            if (!($a)) {
                if ($Error) {
                    Write-Error "Copy-FilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | ERROR: nothing copied for: '$Source' $($Error[0].Exception.Message)"
                    $Global:Error.RemoveAt(1)
                }
                else {
                    Write-Output "Copy-FilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | Nothing to be copied, the source folder '$Source' is empty."
                }
            }
        }
    }
}

Function Get-QuotaHomeDriveHC {
    <#
    .SYNOPSIS
        Retrieve home drive quotas.

    .DESCRIPTION
        Retrieve home drive quota details (usage size, type, ..) from specific servers. The generated
        output can be used to asses which hard quota size limits can be applied, without blocking
        a user.

        This function only reports data and makes no changes to the file system nor to the quota limits.
        For this reason it can be used prior to implementing the script 'Home drive quota' which does make
        changes to the home folder quotas by setting hard quota size limits.

    .PARAMETER ComputerName
        Specifies the computers on which the command runs.

    .EXAMPLE
        Retrieve all home drives and their quotas from two servers.

        'Server1', 'Server2' | Get-QuotaHomeDriveHC -Verbose |
        Select-Object -ExcludeProperty RunSpaceId, userSid, UserAccount -Property @{N='QuotaLimitGB';E={$_.QuotaLimit/1GB}},
        @{N='QuotaUsedGB';E={[Math]::Round(($_.QuotaUsed/1GB), 2)}}, * | Out-GridView
    #>

    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [String[]]$ComputerName
    )

    Process {
        foreach ($C in $ComputerName) {
            Invoke-Command -ComputerName $C -ScriptBlock {
                $VerbosePreference = $Using:VerbosePreference

                $FQM = New-Object -ComObject Fsrm.FsrmQuotaManager

                $HomeDrives = Get-ChildItem -Path ((Get-WmiObject -Class Win32_Share -Filter "Name = 'HOME'").Path) |
                Where-Object { $_.PSIsContainer }

                foreach ($H in $HomeDrives) {
                    Try {
                        Write-Verbose "''$env:COMPUTERNAME'' Get quota '$($H.FullName)'"
                        $FQM.getquota($H.FullName)
                    }
                    Catch {
                        Write-Error "Failed retrieving quota on '$env:COMPUTERNAME' for '$($H.FullName)': No quota set"
                    }
                }
            }
        }
    }
}

Function Get-QuotaLimitHC {
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager

        Try {
            Write-Verbose "'$Path' Get quota limit"
            $FQTM.GetQuota($Path)
        }
        Catch {
            throw 'No quota management configured'
        }
    }
    Catch {
        throw "Failed retrieving the quota limit for '$Path' on '$env:COMPUTERNAME': $_"
    }
}

Function Get-QuotaWarningHC {
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path,
        [ValidateRange(1, 100)]
        [Int]$Threshold
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager
        Try {
            $Quota = $FQTM.GetQuota($Path)
        }
        Catch {
            throw 'No quota management configured'
        }

        if ($Threshold) {
            Write-Verbose "'$Path' Get quota warning with threshold '$Threshold'"
            if ($Quota.Thresholds -contains $Threshold) {
                [PSCustomObject]@{
                    Threshold = $Threshold
                    Actions   = $Quota.EnumThresholdActions($Threshold) | Select-Object *
                }
            }
        }
        else {
            Write-Verbose "'$Path' Get quota warnings"
            ForEach ($N in $Quota.Thresholds) {
                [PSCustomObject]@{
                    Threshold = $N
                    Actions   = $Quota.EnumThresholdActions($N) | Select-Object *
                }
            }
        }
    }
    Catch {
        throw "Failed retrieving the quota warning for '$Path' on '$env:COMPUTERNAME': $_"
    }
}

Function Get-ValueFromArrayHC {
    <#
        .SYNOPSIS
            Search for a specific string in an array and return its values.

        .DESCRIPTION
            Search for a specific string in an array and return its values. We assume that the following format is used, comma separated:
            Name: value1, value2, value3
            Name: value4, value5, ...

            If found we return the values, if not found we don't return anything.

        .PARAMETER Delimiter
            The delimiter used for splitting the arguments

        .PARAMETER Name
            Get all the values matching the patter 'Name:'

        .PARAMETER Exclude
            Get all the values except those defined in Exclude

        .PARAMETER Array
            The array to search

        .EXAMPLE
            $File = @(
            "MailTo: James@hc.com  , Bob@hc.com                      "
            "MailTo: Mike@hc.com  , Chuck@hc.com                     "
            "MailTo:    Jake@hc.com,                                 "
            "                                                        "
            "OUs: OU=Users,OU=CountryA,OU=Region,DC=domain,DC=com    "
            "OUs: OU=Users,OU=CountryB,OU=Region,DC=domain,DC=com    "
            )
            $File | Get-ValueFromArrayHC mailto -Delimiter ','

            Returns:
            James@hc.com
            Bob@hc.com
            Mike@hc.com
            Chuck@hc.com
            Jake@hc.com

        .EXAMPLE
            $File = @(
            "MailTo: James@hc.com  , Bob@hc.com                      "
            "MailTo: Mike@hc.com  , Chuck@hc.com                     "
            "MailTo:    Jake@hc.com,                                 "
            "                                                        "
            "OUs: OU=Users,OU=CountryA,OU=Region,DC=domain,DC=com    "
            "OUs: OU=Users,OU=CountryB,OU=Region,DC=domain,DC=com    "
            )
            $File | Get-ValueFromArrayHC OUs

            Returns:
            OU=Users,OU=CountryA,OU=Region,DC=domain,DC=com
            OU=Users,OU=CountryB,OU=Region,DC=domain,DC=com

        .EXAMPLE
            $File = @(
            "MailTo: James@hc.com  , Bob@hc.com   "
            "Switches: /MIR /Z /MT                "
            "                                     "
            "MaxThreads: 3                        "
            "                                     "
            "Apples                               "
            "Bananas                              "
            )
            $File | Get-ValueFromArrayHC -Exclude MaxThreads, MailTo, Switches

            Returns:
            Apples
            Bananas

        .EXAMPLE
            $File = @(
            "MailTo: James@hc.com  , Bob@hc.com   "
            "Switches: /MIR /Z /MT                "
            "                                     "
            "MaxThreads: 3                        "
            "                                     "
            "Apples; Kiwi                         "
            "Bananas ; Orange                     "
            )
            $File | Get-ValueFromArrayHC -Exclude MaxThreads, MailTo, Switches -Delimiter ';'

            Returns:
            Apples
            Kiwi
            Bananas
            Orange

        .EXAMPLE
            $File = @(
            "MailTo: James@hc.com  , Bob@hc.com                      "
            "Switches: /MIR /Z /MT                                   "
            "                                                        "
            )
            $File | Get-ValueFromArrayHC Switches -Delimiter ' '

            Returns:
            /MIR
            /Z
            /MT
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, ParameterSetName = 'A')]
        [Parameter(Mandatory = $true, ParameterSetName = 'B', Position = 0)]
        [String]$Name,
        [Parameter(Mandatory = $true, ParameterSetName = 'A')]
        [Parameter(Mandatory = $false, ParameterSetName = 'B')]
        [String[]]$Exclude,
        [Parameter(Mandatory = $true, ParameterSetName = 'A', ValueFromPipeline)]
        [Parameter(Mandatory = $true, ParameterSetName = 'B', ValueFromPipeline)]
        [String[]]$Array,
        [String]$Delimiter
    )

    Begin {
        $Result = @()

        if ($Exclude) {
            $Pattern = ($Exclude | ForEach-Object { '^' + $_ + ':' }) -join '|'
            Write-Verbose "Exclude '$Pattern' with delimiter '$Delimiter'"
        }
        else {
            $Pattern = "$Name\:\s*\S+\s*"
            Write-Verbose "Match '$Pattern' with delimiter '$Delimiter'"
        }
    }

    Process {
        if ($Exclude) {
            $Result += $_ | Where-Object { $_ -notMatch $Pattern }
        }
        else {
            $Result += $_ | Where-Object { $_ -match $Pattern } | ForEach-Object {
                $_ -replace "^$Name`:"
            }
        }
    }

    End {
        if ($Delimiter) {
            $Result = $Result -split "$Delimiter\s*"
        }

        $Result | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object {
            $_
            Write-Verbose "'$_'"
        }
    }
}

Function Invoke-ExternalCommandHC {
    <#
    .SYNOPSIS
        Execute an executable

    .DESCRIPTION
        Execute an executable that is not native to PowerShell. It is convenient to have this in a function,
        it allows for mocking the executable with Pester, which is otherwise not possible.

    .PARAMETER Command
        This is the command itself or the path to the executable

    .PARAMETER Argument
        The arguments used to launch the application

    .EXAMPLE
        Launch the TreeSizePro executable with its arguments.

        $TreeSizePro = 'T:\Test\TreeSize Professional\TreeSize.exe'
        $Arguments = '/NOGUI /DATE /EXPAND 4 /SIZEUNIT 3 /SORTTYPE 0 "T:\"'
        $ArgList = ConvertTo-ArgumentHC -String $Arguments -Path 'T:\Test\Log_Test\Report 15.xlsx' -SheetName Data
        Invoke-ExternalCommandHC $TreeSizePro $ArgList
    #>

    Param (
        [Parameter(Mandatory)]
        [ValidateScript( { Test-Path -LiteralPath $_ -Type Leaf })]
        [String]$Command,
        [Parameter(Mandatory)]
        [String[]]$Argument
    )

    $Global:LASTEXITCODE = 0

    $Result = Start-Process $Command -ArgumentList $Argument -Wait

    if ($LASTEXITCODE -ne 0) {
        Throw "FAiled executing command '$Command' with argument '$Argument': $LASTEXITCODE"
    }

    $Result
}

Function New-TextFileHC {
    <#
        .SYNOPSIS
            Create a new text file based on another text file

        .DESCRIPTION
            Read a text file and replace a specific line of text with
            other lines of text

        .PARAMETER InputPath
            Path to the original text file

        .PARAMETER ReplaceLine
            The line of text in the file that needs to be replaced

        .PARAMETER NewLine
            The new line(s) of text to add in the file, at the same location
            as the old line of text

        .PARAMETER NewFilePath
            Path where the new text file will be saved
    #>

    Param (
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path -Path $_ })]
        [String]$InputPath,
        [Parameter(Mandatory)]
        [String]$ReplaceLine,
        [Parameter(Mandatory)]
        [String[]]$NewLine,
        [Parameter(Mandatory)]
        [String]$NewFilePath,
        [Switch]$Overwrite
    )

    try {
        $fileContent = Get-Content -Path $InputPath

        #region Get specific field in text file
        $lineNumber = for ($i = 0; $i -lt $fileContent.Count; $i++) {
            if ($fileContent[$i] -eq $ReplaceLine) {
                $i
            }
        }

        if ($lineNumber.count -gt 1) {
            throw "Multiple lines found with the text '$ReplaceLine'"
        }
        #endregion

        #region Create new file
        try {
            $params = @{
                Path        = $NewFilePath
                Encoding    = 'utf8'
                ErrorAction = 'Stop'
            }

            if (-not $Overwrite) {
                $params.NoClobber = $true
            }

            (
                $fileContent[0..($lineNumber - 1)] +
                $NewLine +
                $fileContent[(($lineNumber + 1)..$fileContent.Length)]
            ) |
            Out-File @params
        }
        catch {
            throw "Failed to create file '$NewFilePath': $_"
        }
        #endregion
    }
    catch {
        throw "Failed creating a new text file: $_"
    }
}

Function Remove-ImportExcelHeaderProblemOnEmptySheetHC {
    <#
    .SYNOPSIS
        Remove the objects returned from Import-Excel on an empty worksheet.

    .DESCRIPTION
        This function is a fix for a bug in Import-Excel. In case an empty
        worksheet is read by the function Import-Excel, it will generate two
        objects. One object with values exactly the same as the header row
        titles and one object full of null values.

        This function simply filters out these unwanted objects.

    .EXAMPLE
        @(
            [PSCustomObject]@{
                Header1 = 'Header1'
                Header2 = 'Header2'
                Header3 = 'Header3'
            }
            [PSCustomObject]@{
                Header1 = $null
                Header2 = $null
                Header3 = $null
            }
            [PSCustomObject]@{
                Header1 = $null
                Header2 = 'Valid data'
                Header3 = $null
            }
        ) | Remove-ImportExcelHeaderProblemOnEmptySheetHC

        Returns one object with 'Valid data'.
    #>

    Param (
        [Parameter(ValueFromPipeline)]
        $Objects
    )

    Process {
        foreach ($O in $Objects) {
            $PropertiesCount = @($O.PSObject.Properties).Count
            $NullValueCount = @($O.PSObject.Properties | Where-Object { -not $_.Value }).Count

            if ($PropertiesCount -eq $NullValueCount) {
                # all properties are null
                # object not returned
            }
            elseif (
                ($EqualValue = $O.PSObject.Properties | Where-Object { $_.Name -eq $_.Value }) -and
                ($EqualValue.Count -eq $PropertiesCount)
            ) {
                # all property values are the same as their property name
                # object not returned
            }
            else {
                $O
            }
        }
    }
}

Function Remove-LeaverHomeDriveHC {
    <#
.SYNOPSIS
    Remove the home drive folder of a leaver.

.DESCRIPTION
    Folders of users that left the company are stored in a fixed location. These
    folders stay there for a couple of months so they are stored on the backup
    media.

    After the folder data is stored in de backups this function can be used to
    delete the folder permanently from the system.

.EXAMPLE
    Remove-LeaverHomeDriveHC -SamAccountName 'deleusd', 'dverhuls'
    Removes the folders 'deleus' and 'dverhuls' and all their data from the
    leaver folder.
#>

    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [String[]]$SamAccountName,
        [String]$ComputerName = 'DEUSFFRAN0031'
    )

    Invoke-Command  -ComputerName $ComputerName -ScriptBlock {
        $VerbosePreference = $Using:VerbosePreference

        foreach ($S in $Using:SamAccountName) {
            $Path = "E:\HOME\000 LEAVERS\$S"

            if (Test-Path -Path $Path -PathType Container) {
                try {
                    @(Get-SmbOpenFile -IncludeHidden).Where( { $_.Path -like "$Path\*" }) |
                    Close-SmbOpenFile -Force

                    Remove-Item -Path $Path -Recurse -Force -EA Stop

                    Write-Host "Removed folder '$Path'" -ForegroundColor Green
                }
                catch {
                    Write-Error "Failed removing home drive '$Path': $_"
                }
            }
            else {
                Write-Warning "Folder '$Path' not found"
            }
        }
    }
}

Function Remove-QuotaLimitHC {
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager
        $fqtmTemplate = New-Object -ComObject Fsrm.FsrmQuotaTemplateManager

        Try {
            $Quota = $FQTM.GetQuota($Path)
            $Quota.Description = $null
            Write-Verbose "'$Path' Remove quota limit and restore default quota template"
            $Quota.ApplyTemplate($Quota.SourceTemplateName)

            # Set quota template flag, to avoid 'Disabled quota' checkbox
            $QuotaTemplate = $fqtmTemplate.GetTemplate($Quota.SourceTemplateName)
            $Quota.QuotaFlags = $QuotaTemplate.QuotaFlags

            $Quota.Commit()
        }
        Catch {
            throw 'No quota management configured'
        }
    }
    Catch {
        throw "Failed removing quota for path '$Path' on '$env:COMPUTERNAME': $_"
    }
}

Function Remove-QuotaWarningHC {
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path,
        [ValidateRange(1, 100)]
        [Int]$Threshold
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager

        Try {
            $Quota = $FQTM.GetQuota($Path)
        }
        Catch {
            throw 'No quota management configured'
        }

        if ($Threshold) {
            Write-Verbose "'$Path' Remove quota warning with threshold '$Threshold'"
            $Quota.DeleteThreshold($Threshold)
        }
        else {
            Write-Verbose "'$Path' Remove all quota warnings"
            $Quota.Thresholds | ForEach-Object { $Quota.DeleteThreshold($_) }
        }
        $Quota.Commit()
    }
    Catch {
        throw "Failed removing the quota warning for '$Path' on '$env:COMPUTERNAME': $_"
    }
}

Function Set-QuotaLimitHC {
    [CmdLetBinding()]
    Param (
        [ValidateScript( {
                if (Test-Path -LiteralPath $_ -PathType Container) {
                    $true
                }
                else {
                    throw "Couldn't find the path '$_' on '$env:COMPUTERNAME'"
                }
            })]
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [Double]$Limit
    )

    Try {
        $FQTM = New-Object -ComObject Fsrm.FsrmQuotaManager
        Try {
            $Quota = $FQTM.GetQuota($Path)
        }
        Catch {
            throw 'No quota management configured'
        }
        Write-Verbose "'$Path' Set quota to '$Limit' hard"
        $Quota.QuotaLimit = $Limit
        $Quota.Description = "Custom limit set by PowerShell based on AD group membership"
        $Quota.QuotaFlags = 256 # hard
        $Quota.Commit()
    }
    Catch {
        throw "Failed setting quota limit '$Limit' for '$Path' on '$env:COMPUTERNAME': $_"
    }
}

Function Lock-FileHC {
    $FileLockOn = '\\grouphc.net\bnl\DEPARTMENTS\Brussels\CBR\SHARE\Test\Monitor\Folder A\File (2).rtf'
    $file = [System.io.File]::Open($FileLockOn, 'append', 'Write', 'None')
    $enc = [system.Text.Encoding]::UTF8
    $msg = "This is a test"
    $data = $enc.GetBytes($msg)
    $file.write($data, 0, $data.length)
    $file.Close()
}
Function Move-ToArchiveHC {
    <#
        .SYNOPSIS
            Moves files to folders based on their creation date.

        .DESCRIPTION
            Moves files to folders, where the destination folder names are automatically created based on the file's creation date (year/month). This is useful in situation where files need to be archived by year for example. When a file already exists on the destination it will be overwritten. When a file is in use by another process, we can't move it so we only report that it's in use, no error is thrown.

        .PARAMETER Source
            The source folder where we will pick up the files to move them to the destination folder. This folder is only used to pick up files on the root directory, so not recursively.

        .PARAMETER Destination
            The destination folder where the files will be moved to. When left empty, the files will be moved to sub folders that will be created in the source folder.

        .PARAMETER Structure
            The folder structure that will be used on the destination. The files will be moved based on their creation date. The default value is 'Year-Month'. Valid options are:

            'Year'
            C\SourceFolder\2014
            C\SourceFolder\2014\File december.txt

            'Year\Month'
            C\SourceFolder\2014
            C\SourceFolder\2014\12
            C\SourceFolder\2014\12\File december.txt

            'Year-Month'
            C\SourceFolder\2014-12
            C\SourceFolder\2014-12\File december.txt

            'YYYYMM'
            C\SourceFolder\201504
            C\SourceFolder\201504\File.txt

        .PARAMETER OlderThan
            This is a filter to only archive files that are older than x days/months/years, where 'x' is defined by the parameter 'Quantity'. When 'OlderThan' and 'Quantity' are not used, all files will be moved an no filtering will take place. Valid options are:
            'Day'
            'Month'
            'Year'

        .PARAMETER Quantity
            Quantity defines the number of days/months/years defined for 'OlderThan'. Valid options are only numbers.

            -OlderThan Day -Quantity '3'    > All files older than 3 days will be moved
            -OlderThan Month -Quantity '1'  > All files older than 1 month will be moved (all files older than this month will be moved)
            <blanc>                         > All files will be moved, regardless of their creation date

        .EXAMPLE
            Move-ToArchiveHC -Source 'T:\Truck movements' -Verbose
            Moves all files based on their creation date from the folder 'T:\Truck movements' to the folders:
            'T:\Truck movements\2014-01\File Jan 2014.txt', 'T:\Truck movements\2014-02\File Feb 2014.txt', ..

        .EXAMPLE
            Move-ToArchiveHC -Source 'T:\GPS' -Destination 'C:\Archive' -Structure Year\Month -OlderThan Day -Quantity '3' -Verbose
            Moves all files older than 3 days, based on their creation date, from the folder 'T:\GPS' to the folders:
            'C:\Archive\2014\01\2014-01-01.xml', 'C:\Archive\2014\01\2014-01-02.xml', 'C:\Archive\2014\01\2014-01-03.xml' ..
    #>

    [CmdletBinding(SupportsShouldProcess = $True, DefaultParameterSetName = 'A')]
    Param (
        [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'A')]
        [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'B')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { Test-Path $_ -PathType Container })]
        [String]$Source,
        [parameter(Mandatory = $false, Position = 1, ParameterSetName = 'A')]
        [parameter(Mandatory = $false, Position = 1, ParameterSetName = 'B')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { Test-Path $_ -PathType Container })]
        [String]$Destination = $Source,
        [parameter(Mandatory = $false, ParameterSetName = 'A')]
        [parameter(Mandatory = $false, ParameterSetName = 'B')]
        [ValidateSet('Year', 'Year\Month', 'Year-Month', 'YYYYMM')]
        [String]$Structure = 'Year-Month',
        [parameter(Mandatory = $true, ParameterSetName = 'B')]
        [ValidateSet('Day', 'Month', 'Year')]
        [String]$OlderThan,
        [parameter(Mandatory = $true, ParameterSetName = 'B')]
        [Int]$Quantity
    )

    Begin {
        $Today = Get-Date

        Switch ($OlderThan) {
            'Day' {
                Filter Select-Stuff {
                    Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                    if ($_.CreationTime.Date.ToString('yyyyMMdd') -le $(($Today.AddDays( - $Quantity)).Date.ToString('yyyyMMdd'))) {
                        Write-Output $_
                    }
                }
            }
            'Month' {
                Filter Select-Stuff {
                    Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                    if ($_.CreationTime.Date.ToString('yyyyMM') -le $(($Today.AddMonths( - $Quantity)).Date.ToString('yyyyMM'))) {
                        Write-Output $_
                    }
                }
            }
            'Year' {
                Filter Select-Stuff {
                    Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                    if ($_.CreationTime.Date.ToString('yyyy') -le $(($Today.AddYears( - $Quantity)).Date.ToString('yyyy'))) {
                        Write-Output $_
                    }
                }
            }
            Default {
                Filter Select-Stuff {
                    Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                    Write-Output $_
                }
            }
        }

        Write-Output @"
    ComputerName: $Env:COMPUTERNAME
    Source:       $Source
    Destination:  $Destination
    Structure:    $Structure
    OlderThan:    $OlderThan
    Quantity:     $Quantity
    Date:         $($Today.ToString('dd/MM/yyyy hh:mm:ss'))

    Moved file:
"@
    }

    Process {
        $File = $null

        Get-ChildItem $Source -File | Select-Stuff | ForEach-Object {
            $File = $_

            $ChildPath = Switch ($Structure) {
                'Year' { [String]$File.CreationTime.Year }
                'Year\Month' { [String]$File.CreationTime.Year + '\' + $File.CreationTime.ToString('MM') }
                'Year-Month' { [String]$File.CreationTime.Year + '-' + $File.CreationTime.ToString('MM') }
                'YYYYMM' { [String]$File.CreationTime.Year + $File.CreationTime.ToString('MM') }
            }
            $Target = Join-Path -Path $Destination -ChildPath $ChildPath

            Try {
                $null = New-Item $Target -Type Directory -EA Ignore
                Move-Item -Path $File.FullName -Destination $Target -EA Stop
                Write-Output "- '$File' > '$ChildPath'"
            }
            Catch {
                Switch ($_) {
                    { $_ -match 'cannot access the file because it is being used by another process' } {
                        Write-Output "- '$File' WARNING $_"
                        $Global:Error.RemoveAt(0)
                        break
                    }
                    { $_ -match 'file already exists' } {
                        Move-Item -Path $File.FullName -Destination $Target -Force
                        Write-Output "- '$File' WARNING File already existed on the destination but has now been overwritten"
                        $Global:Error.RemoveAt(0)
                        break
                    }
                    default {
                        Write-Error "Error moving file '$($File.FullName)': $_"
                        $Global:Error.RemoveAt(1)
                        Write-Output "- '$File' ERROR $_"
                    }
                }
            }
        }

        if (-not $File) {
            Write-Output '- INFO No files found that match the filter, nothing moved'
        }
    }
}
Function New-FolderHC {
    <#
    .SYNOPSIS
        Create a new folder.

    .DESCRIPTION
        Create a new folder after joining the Path and the ChildPath into a single path. The output will be the full path name of the folder.

    .PARAMETER Path
        Specifies the main path to which the child-path is appended. This needs to be an already existing path.

    .PARAMETER ChildPath
        Specifies the elements to append to the value of Path. This is the folder that will be created in the already existing Path.

    .EXAMPLE
        New-FolderHC -Path "\\$env:COMPUTERNAME\s$\Test\Log_Test" -ChildPath 'Reports\Finance'
        \\$env:COMPUTERNAME\s$\Test\Log_Test\Reports\Finance

        Creates the path "\\$env:COMPUTERNAME\s$\Test\Log_Test\Reports\Finance" and output the FullName of the path.
    #>

    [CmdLetBinding()]
    [OutputType([System.IO.FileSystemInfo])]
    Param (
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [String]$ChildPath
    )

    Try {
        $FullPath = Join-Path -Path $Path -ChildPath $ChildPath -EA Stop

        if (-not (Test-Path $Path -PathType Container)) {
            throw "Path '$Path' not found"
        }

        if (-not (Test-Path -PathType Container $FullPath)) {
            $null = New-Item -ItemType Directory $FullPath -EA Stop
            Write-Verbose "Created folder '$FullPath'"
        }

        Get-Item -LiteralPath $FullPath
    }
    Catch {
        throw "Failed creating folder '$ChildPath': $_"
    }
}
Function Remove-OldFilesHC {
    <#
    .SYNOPSIS
        Function to delete files older than x days and delete empty folders if requested.

    .DESCRIPTION
        Remove files older than x days in all subfolders and write success and failure actions to the log file.
        By default, empty folders will be left behind and not deleted.

    .PARAMETER Target
        The path that will be recursively scanned for old files.

    .PARAMETER OlderThanDays
        Filter for age of file, entered in days. Use '0' to delete all files in the subfolder.

    .PARAMETER CleanFolders
        If set to '$True' all empty folders will be removed, regardless of the creation date. Default behavior is to leave empty folders behind/untouched.

    .EXAMPLE
        Remove-OldFilesHC -Target '\\grouphc.net\bnl\DEPARTMENTS\Brussels\Archive' -OlderThanDays 10
        Deletes all files older than 10 days in the 'Archive' folder and all of its subfolders.

    .EXAMPLE
        Remove-OldFilesHC -Target 'E:\Departments\Log' -OlderThanDays 5 -CleanFolders 1
        Deletes all files older than 5 days in the 'Log' folder and all of its subfolders. Afterwards all empty folders will be deleted to.
 #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateScript( {
                if (Test-Path -LiteralPath $_) { $true }
                else { throw "Path '$_' not found on '$env:COMPUTERNAME'" }
            })]
        [String]$Target,
        [Parameter(Mandatory)]
        [Int]$OlderThanDays,
        [Parameter()]
        [Switch]$CleanFolders
    )
    Begin {
        $compareDate = (Get-Date).AddDays(-$OlderThanDays)

        Filter Select-FilesHC {
            if ($_.CreationTime -lt $compareDate) {
                Write-Output $_
            }
        }

        Function Remove-ItemHC {
            Param (
                [Parameter(Mandatory)]
                $Item
            )
            try {
                Write-Verbose "Remove Item '$($Item.FullName)' created at $($Item.CreationTime)"
                $Item | Remove-Item  -Recurse -Force -EA Stop
                "Remove-OldFilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | REMOVED: $($Item.FullName)"
            }
            Catch {
                "Remove-OldFilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | FAILED: $($Item.FullName) : $_"
            }
        }
    }
    Process {
        $fileToRemove = $false

        if (Test-Path $Target -PathType Container) {
            Write-Verbose "Remove files older than '$compareDate' or $OlderThanDays days"

            Get-ChildItem -LiteralPath $Target -Recurse -File | Select-FilesHC |
            ForEach-Object {
                $fileToRemove = $True
                Remove-ItemHC -Item $_
            }

            if ($CleanFolders) {
                Write-Verbose 'Remove empty folders'

                $failedFolderRemoval = @()

                while ($emptyFolders = Get-ChildItem -LiteralPath $Target -Recurse -Directory |
                    Where-Object { ($_.GetFileSystemInfos().Count -eq 0) -and ($failedFolderRemoval -notContains $_.FullName) }) {
                    $emptyFolders | ForEach-Object {
                        $fileToRemove = $True
                        Remove-ItemHC -Item $_

                        if (Test-Path -LiteralPath $_.FullName) {
                            $failedFolderRemoval += $_.FullName
                        }
                    }
                }
            }
        }

        else {
            Get-ChildItem -Path $Target -File | Select-FilesHC |
            ForEach-Object {
                $fileToRemove = $True
                Remove-ItemHC -Item $_
            }
        }
        if (!($fileToRemove)) {
            Write-Output "Remove-OldFilesHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | Nothing to be removed, no files older than '$OlderThanDays' days on '$Target'."
        }
    }
}
Function Search-ScriptsHC {
    <#
    .SYNOPSIS
        Searches for a string pattern of text in specific files.

    .DESCRIPTION
        Searches for a string pattern of text in specific files like script modules and script production files. Calculates how many times it's used and represents it in color.

    .PARAMETER Path
        The parent folder where we recursively look for files.

    .PARAMETER Pattern
        The string pattern to look for.

    .PARAMETER OpenFile
        Opens the found files in the PowerShell ISE when it's a PowerShell file or in Notepad++

    .PARAMETER Filter
        The include used with Get-ChildItem on which the selection is based.
        Can be file extensions like
        '*.ps1','*.psm1'
        '*.csv','*.txt'

    .EXAMPLE
         Search-FunctionHC Send-MailHC
         Lists all the scripts and modules containing the string 'Send-MailHC':
         Count Name                                    Path
         ----- ----                                    ----
             4 AD Computers OS.ps1                     T:\Prod\AD Reports\AD Computers OS.ps1
             4 AD Group members managers.ps1           T:\Prod\AD Reports\AD Group members managers.ps1
             1 AD HR User list.ps1                     T:\Prod\AD Reports\AD HR User list.ps1

    .EXAMPLE
         Search-FunctionHC Search-FunctionHC -OpenInIse
         Lists all the scripts and modules containing the string 'Search-FunctionHC' and opens them in the PowerShell ISE.

    .EXAMPLE
        Search-ScriptsHC -Pattern Appels -Path 'T:\Input' -Include '*.csv','*.txt', '*.ps1'
        Find all input files that contain the word Appel

    .EXAMPLE
        Search-ScriptsHC -Pattern Send-MailAuthenticatedHC -OpenFile
        Open all files where we use this CmdLet
    #>

    [CmdLetBinding()]
    Param (
        [String]$Pattern,
        [String[]]$Path = @(
            'T:\Prod',
            'T:\Input',
            'C:\Program Files\WindowsPowerShell\Modules',
            'C:\Program Files\PowerShell\Modules'
        ),
        [String[]]$Include = @('*.ps1', '*.psm1', '*.psd', '*.json', '*.csv', '*.txt'),
        [String]$NotepadPlusPlus = 'C:\Program Files\Notepad++\notepad++.exe',
        [Switch]$OpenFile
    )

    Begin {
        Try {
            $TotalMatches = 0

            if (($OpenFile) -and (-not (Test-Path $NotepadPlusPlus))) {
                throw "Notepad++ not installed"
            }

            $ExecutableProgram = switch ($Host.Name) {
                'Visual Studio Code Host' { 'code'; Break }
                'Windows PowerShell ISE Host' { 'psEdit' ; Break }
                Default { $NotepadPlusPlus }
            }
        }
        Catch {
            throw "Failed searching for pattern '$Pattern': $_"
        }
    }

    Process {
        Try {
            $Params = @{
                File        = $true
                Recurse     = $true
                LiteralPath = $Path
                Include     = $Include
                ErrorAction = 'Stop'
            }
            Get-ChildItem @Params |
            Where-Object { $Include -Contains "*$($_.Extension)" } | ForEach-Object {
                $MatchCount = (
                    Select-String -LiteralPath $_ -Pattern $Pattern -AllMatches -SimpleMatch
                ).Count

                if ($MatchCount -ge 1) {
                    $_ | Select-Object @{N = 'Count'; E = { $MatchCount } }, Name, FullName

                    if ($OpenFile) {
                        if ($_.Extension -match '.ps1|.psm1|.psd|.json') {
                            & $ExecutableProgram $_.FullName
                        }
                        else {
                            & $NotepadPlusPlus $_.FullName
                        }
                    }

                    $TotalMatches += $MatchCount
                }
            }

            Write-Verbose "Found '$TotalMatches' matches for pattern '$Pattern' in '$Path'"
        }
        Catch {
            throw "Failed searching for pattern '$Pattern': $_"
        }
    }
}
Function Test-WriteToFolderHC {
    <#
    .SYNOPSIS
        Function to check if we have write permissions in a folder.

    .DESCRIPTION
        This function checks if we are able to write (files) in a folder by writing a random file and checking the result. On success we return $true and on failure $false.

    .PARAMETER Path
        The path that will be checked.

    .EXAMPLE
        Test-WriteToFolderHC "\\domain.net\share"
        Checks if we have write permissions on "\\domain.net\share" and returns $true if we can write and $false if we can't.

    .EXAMPLE
        "C:\Program Files\WindowsPowerShell" | Test-WriteToFolderHC
        Checks if we have write permissions on "C:\Program Files\WindowsPowerShell" and returns $true if we can write and $false if we can't.

    .EXAMPLE
        "L:\Scheduled Task\Auto_Clean", "L:\Scheduled Task" | Test-WriteToFolderHC
        Checks if we have write permissions on both folders and returns $true or $false for each one.

    .EXAMPLE
        Test-WriteToFolderHC "L:\Scheduled Task\Auto_Clean", "L:\Scheduled Task"
        Checks if we have write permissions on both folders and returns $true or $false for each one.
    #>

    [CmdletBinding(SupportsShouldProcess = $True)]
    Param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { Test-Path $_ -PathType Container })]
        [String[]] $Path
    )

    Process {
        foreach ($_ in $Path) {
            try {
                Write-Verbose "Function Can-WriteToFolder: Write a new file to the folder '$_'"
                $TestPath = Join-Path $_ $(Get-Random)
                New-Item -Path $TestPath -ItemType File -ErrorAction Stop > $null
                Write-Verbose "Return TRUE for '$_'"
                $true
            }
            catch {
                Write-Verbose "Function Can-WriteToFolder: Catch return FALSE for '$_'"
                $false
            }
            finally {
                Write-Verbose "Function Can-WriteToFolder: Remove the random item '$TestPath'"
                Remove-Item $TestPath -ErrorAction SilentlyContinue
            }

        }
    }
}
Function Watch-FolderForChangesHC {
    <#
    .SYNOPSIS
        Monitor a folder for changes and trigger actions when needed.

    .DESCRIPTION
        Monitor the parent folder for changes (Changed, Deleted, Created,
        Renamed) of files and/or folders.When a change or event is triggered,
        the action scriptblock is called to execute code.

    .PARAMETER Path
        Specifies the parent folder to monitor.

    .PARAMETER Filter
        What type of files/folders to monitor.

    .PARAMETER NotifyFilter
        Defines when an action is triggered:
        - 'FileName' monitors files only
        - 'DirectoryName' monitors folders only
        Options can be combined too (ex. 'FileName, LastWrite, DirectoryName').
        By default, everything is monitored.

    .PARAMETER Recurse
        Monitor subfolders.

    .PARAMETER CreatedAction
        Action to take when a new file/folder is created.

        Variables available in the 'Global' scope that can be used in the
        scriptblock are:
        - $Event: All event properties
        - $EventName: Name of the file/folder
        - $EventChangeType: Created, Deleted, ...

    .PARAMETER DeletedAction
        Action to take when a file/folder is removed.

    .PARAMETER ChangedAction
        Action to take when a file/folder is changed.

    .PARAMETER RenamedAction
        Action to take when a file/folder is renamed.

    .PARAMETER Timeout
        Specifies the maximum time in seconds to monitor a folder. When the
        maximum seconds are reached the function exits, depending on the
        parameter 'EndlessLoop' or continues running but will restart the
        Watcher service.

        Restarting the Watcher service is required to free up memory allocation.
        This will ensure that events are still triggering when a network
        connection issue occurred.

    .PARAMETER EndlessLoop
        Defines if the function runs forever or needs to stop running after a
        specified amount of secondes, defined in 'TimeOut'.

    .PARAMETER LogFile
        Finally, because a server typically runs unattended, it's convenient to
        be able to append all output of the event handlers to a log file.
        Start-FileSystemWatcher can do that for you when you supply the
        -LogFile <string> argument. When an event handler action produces
        output, it's written to output and append to the log fie. You can, of
        course, also have your event handlers use Write-Verbose or Write-Host,
        which are not written to the log file.

    .EXAMPLE
        Monitor the creation and deletion of files in the folder 'T:\Test'

        $params = @{
            Path          = 'T:\Test'
            NotifyFilter  = 'FileName'
            CreatedAction = {
                Write-Host "Folder '$EventName' has been '$EventChangeType'" -ForegroundColor Yellow
                # $Event Contains all properties
            }
            DeletedAction = {
                Write-Host "Folder '$EventName' has been '$EventChangeType'" -ForegroundColor Green
            }
            EndlessLoop   = $true
            Timeout       = 30
        }
        Watch-FolderForChangesHC @params -Verbose
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateScript( { Test-Path -Path $_ -PathType 'Container' })]
        [String]$Path,
        [Parameter(Position = 1)]
        [String]$Filter = '*.*',
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { ($_ -notMatch 'FileRenamed|FileChanged') })]
        [IO.NotifyFilters]$NotifyFilter = 'FileName, LastWrite, DirectoryName',
        [Switch]$Recurse,
        [Scriptblock]$CreatedAction,
        [Scriptblock]$DeletedAction,
        [ValidateRange(1, 1800)]
        [Int]$Timeout = '300',
        [Boolean]$EndlessLoop = $true,
        [String]$LogFile
    )

    Begin {
        Function Start-ActionHC {
            <#
                .SYNOPSIS
                    The action to execute

                .PARAMETER Action
                    The action to execute. Is one of the script arguments $ChangedAction, $CreateAction, etc.
                    and if needed writes the output to a file
            #>

            Param (
                [Scriptblock]$Action
            )

            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Execute code"

            $output = Invoke-Command $Action

            if ($output) {
                Write-Output $output

                if ($LogFile) {
                    Write-Output $output >> $LogFile
                }
            }
        }

        Function Unregister-EventsHC {
            @('FileCreated', 'FileDeleted').ForEach(
                {
                    Unregister-Event -SourceIdentifier $_ -ErrorAction Ignore
                    Remove-Event -SourceIdentifier $_ -ErrorAction Ignore
                }
            )
        }

        Function Register-EventsHC {
            if ($CreatedAction) {
                Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated
            }
            if ($DeletedAction) {
                Register-ObjectEvent $fsw Deleted -SourceIdentifier FileDeleted
            }
        }

        Function New-Watcher {
            $folderContent = @{}
            $folderContent.Before = Get-ChildItem -Path $Path -Filter $Filter -Recurse:$Recurse

            if ($fsw) {
                $fsw.Dispose() # free up memory
                Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Stop watcher"
            }

            Unregister-EventsHC

            if (-not (Test-Path -Path $Path -PathType 'Container')) {
                throw "Path '$Path' not found"
            }

            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Start watcher"

            [System.IO.FileSystemWatcher]$fsw = New-Object System.IO.FileSystemWatcher $Path, $Filter -Property @{
                IncludeSubdirectories = $Recurse
                InternalBufferSize    = 16384
                NotifyFilter          = [IO.NotifyFilters]$NotifyFilter
            }

            Register-EventsHC

            $fsw.EnableRaisingEvents = $true

            $folderContent.After = Get-ChildItem -Path $Path -Filter $Filter -Recurse:$Recurse

            if ($folderContent.Before -or $folderContent.After) {
                if ($CreatedAction) {
                    $folderContent.After.where(
                        { $folderContent.Before.Name -notContains $_.Name },
                        'First'
                    ).foreach(
                        {
                            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - File created during watcher restart"

                            $EventName = $_.Name
                            $EventChangeType = 'Created'

                            Start-ActionHC $CreatedAction
                        }
                    )
                }
                if ($DeletedAction) {
                    $folderContent.Before.where(
                        { $folderContent.After.Name -notContains $_.Name },
                        'First'
                    ).foreach(
                        {
                            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - File removed during watcher restart"

                            $EventName = $_.Name
                            $EventChangeType = 'Removed'

                            Start-ActionHC $DeletedAction
                        }
                    )
                }
            }

            $fsw
        }
    }

    Process {
        Try {
            if (-not ($CreatedAction -or $DeletedAction)) {
                throw 'At least one scriptblock argument needs to be provided'
            }

            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Monitor Path '$Path' Filter '$Filter' Recurse '$Recurse' for '$Timeout' seconds (timer reset after every event)"

            $fsw = New-Watcher

            do {
                $startDate = Get-Date

                #Global vars for use in the script blocks outside the module
                $Global:event = Wait-Event -Timeout $Timeout

                if ((-not $event) -and (-not $EndlessLoop)) {
                    Write-Verbose 'Stop monitoring'
                    break
                }

                if ($event) {
                    [String]$Global:EventName = $event.SourceEventArgs.Name
                    [System.IO.WatcherChangeTypes]$Global:EventChangeType = $Event.SourceEventArgs.ChangeType

                    Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - EventID '$($event.EventIdentifier)' Type '$EventChangeType' Name '$EventName'"

                    switch ($EventChangeType) {
                        Created { Start-ActionHC $CreatedAction; Break }
                        Deleted { Start-ActionHC $DeletedAction; Break }
                        Default {
                            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - EventID '$($event.EventIdentifier)' Type '$EventChangeType' Name '$EventName'"
                        }
                    }

                    Remove-Event -EventIdentifier $($event.EventIdentifier)

                    Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - EventID '$($event.EventIdentifier)' end"
                }

                #region Workaround for network connection issues

                # when the network connection to the Path is gone
                # new files are no longer detected
                # for this reason the watcher has to be restarted

                if (((Get-Date) - $startDate).TotalSeconds -ge $Timeout) {
                    Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Timeout of $Timeout seconds reached"
                    $fsw = New-Watcher
                }
                #endregion
            } while ($true)
        }
        Catch {
            $M = $_
            $global:error.RemoveAt(0)
            throw "Failed monitoring the folder '$Path' for changes: $M"
        }
        Finally {
            Unregister-EventsHC

            if ($fsw) {
                $fsw.Dispose() # free up memory
            }

            Write-Verbose "$((Get-Date).ToString('HH:mm:ss:fff')) - Stopped monitoring"
        }
    }
}
Function Write-ZipHC {
    <#
        .SYNOPSIS
            Creates ZIP file
    #>
    [cmdletBinding()]
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateScript( { Test-Path -Path $_ })]
        [String[]]$Source,
        [ValidateScript( { Test-Path -Path (Split-Path $_) })]
        [Parameter(Mandatory = $True)]
        [String]$Target
    )

    Begin {
        if (-not (Test-Path "$env:ProgramFiles\7-Zip\7z.exe")) {
            throw "$env:ProgramFiles\Z-Zip\7z.exe needed"
        }
        Set-Alias sz "$env:ProgramFiles\7-Zip\7z.exe"
    }

    Process {
        foreach ($S in $Source) {
            Write-Verbose "Zip source '$Source' to destination '$Target'"
            sz a -t7z -m0=lzma2 -mx=5 $Target $S
        }
    }
}

New-Alias 'bs' -Value 'Backup-ScriptsHC' -Description "Backup module alias"

Export-ModuleMember -Function * -Alias *