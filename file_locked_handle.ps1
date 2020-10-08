#https://stackoverflow.com/questions/958123/powershell-script-to-check-an-application-thats-locking-a-file
#usage: PS C:\tmp> Get-LockingProcess C:\tmp\foo.txt
#/* http://jdhitsolutions.com/blog/powershell/3744/friday-fun-find-file-locking-process-with-powershell/ */
Function Get-LockingProcess {

    [cmdletbinding()]
    Param(
        [Parameter(Position=0, Mandatory=$True,
        HelpMessage="What is the path or filename? You can enter a partial name without wildcards")]
        [Alias("name")]
        [ValidateNotNullorEmpty()]
        [string]$Path
    )

    # Define the path to Handle.exe
    # //$Handle = "G:\Sysinternals\handle.exe"
	
	$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $Handle = "$PSScriptRoot\handle64.exe"

    # //[regex]$matchPattern = "(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\b(\d+)\b)\s+type:\s+(?<Type>\w+)\s+\w+:\s+(?<Path>.*)"
    # //[regex]$matchPattern = "(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\d+)\s+type:\s+(?<Type>\w+)\s+\w+:\s+(?<Path>.*)"
    # (?m) for multiline matching.
    # It must be . (not \.) for user group.
    [regex]$matchPattern = "(?m)^(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\d+)\s+type:\s+(?<Type>\w+)\s+(?<User>.+)\s+\w+:\s+(?<Path>.*)$"

    # skip processing banner
    $data = &$handle -u $path -nobanner
    # join output for multi-line matching
    $data = $data -join "`n"
    $MyMatches = $matchPattern.Matches( $data )

    # //if ($MyMatches.value) {
    if ($MyMatches.count) {

        $MyMatches | foreach {
            [pscustomobject]@{
                FullName = $_.groups["Name"].value
                Name = $_.groups["Name"].value.split(".")[0]
                ID = $_.groups["PID"].value
                Type = $_.groups["Type"].value
                User = $_.groups["User"].value.trim()
                Path = $_.groups["Path"].value
                toString = "pid: $($_.groups["PID"].value), user: $($_.groups["User"].value), image: $($_.groups["Name"].value)"
            } #hashtable
        } #foreach
    } #if data
    else {
        Write-Warning "No matching handles found"
    }
} #end function