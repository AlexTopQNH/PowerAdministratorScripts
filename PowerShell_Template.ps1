<#  
.SYNOPSIS  
    A brief description of the function or script. This keyword can be used
    only once in each topic.

.DESCRIPTION  
    A detailed description of the function or script. This keyword can be
    used only once in each topic.

.PARAMETER  <Parameter-Name>
    The description of a parameter. Add a .PARAMETER keyword for
    each parameter in the function or script syntax.

    Type the parameter name on the same line as the .PARAMETER keyword. 
    Type the parameter description on the lines following the .PARAMETER
    keyword. Windows PowerShell interprets all text between the .PARAMETER
    line and the next keyword or the end of the comment block as part of
    the parameter description. The description can include paragraph breaks.

    The Parameter keywords can appear in any order in the comment block, but
    the function or script syntax determines the order in which the parameters
    (and their descriptions) appear in help topic. To change the order,
    change the syntax.
 
    You can also specify a parameter description by placing a comment in the
    function or script syntax immediately before the parameter variable name.
    If you use both a syntax comment and a Parameter keyword, the description
    associated with the Parameter keyword is used, and the syntax comment is
    ignored. 

.NOTES
    Additional information about the function or script.
    This can be as follows:  
    File Name     : PSTemplate.ps1
    Author        : Jev - @devjevnl
    Version       : 1.0
    Last Modified : 31-10-2016
    Changes       :
                    1.1 - 31-10-2016 - JS - Minor change to the script
                    1.0 - 31-10-2016 - JS - Initial script
        
.LINK
    The name of a related topic. The value appears on the line below
    the .LINE keyword and must be preceded by a comment symbol (#) or
    included in the comment block. 

    Repeat the .LINK keyword for each related topic.

    This content appears in the Related Links section of the help topic.

    The Link keyword content can also include a Uniform Resource Identifier
    (URI) to an online version of the same help topic. The online version 
    opens when you use the Online parameter of Get-Help. The URI must begin
    with "http" or "https"

.EXAMPLE
    A sample command that uses the function or script, optionally followed
    by sample output and a description. Repeat this keyword for each example.


#>

#Script parameters
param ($inifile, $xmlfile)

#region ---------------------------------------- Initialisation ------------------------------------------

# Uncomment should the script require enforcing Windows PowerShell 2.0 or a later
#requires -Version 2.0

<#
    Terminating Error: A serious error during execution that halts the command (or script execution) completely. 
    Examples can include non-existent cmdlets, syntax errors that would prevent a cmdlet from running, or other fatal errors.
    Non-Terminating Error: A non-serious error that allows execution to continue despite the failure. 
    Examples include operational errors such file not found, permissions problems, etc.
#>

<#
    Setting this variable to Stop ensures thath any Non-Terminating errors are handled the same way a Terminating Error is
    Ensuring that the script will stop processing on if an error occurs
#>
$ErrorActionPreference = 'Stop'

# Colors of errors, warnings and notifications
$ColorError = 'Red'
$ColorWarning = 'Yellow'
$ColorAttention = 'Cyan'
$ColorExtraAttention = 'Magenta'

# Log File Info
$LogDir = "D:\beheer\logfiles\SP2013\BeheerTaken"
$LogFileName = "$LogDir\$($(get-item $($MyInvocation.mycommand.path)).basename)-$(get-date -format yyyyMMdd-HHmmss)-$pid.log"

# Import Modules & Snap-ins
# Load Global Functions
. D:\Beheer\Scripts\SP2013\GlobalFunctions.ps1


#endregion ------------------------------------ End Initialisation ---------------------------------------

#region ---------------------------------------- Helper Functions  ---------------------------------------
    <#
        Functions specific to this specific task and script
        which are called from the Main functions should
        reside in this region
    #>
#endregion ------------------------------------- End Helper Functions  -----------------------------------

#region ---------------------------------------- MAIN ----------------------------------------------------
# Always use a function Main() in this framework
Function Main([string[]]$arr)
{

    # Declare variables
    $status = 0

    If ($RunInteractive)
    {
        RDW-Write-Host "Asking user to confirm..."
        Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
		Write-Host -Foregroundcolor $ColorAttention "About to START [YOUR TEXT HERE] on server" -NoNewline
	    Write-Host -Foregroundcolor $ColorExtraAttention "$($env:computerName)"
		Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
        Write-Host -Foregroundcolor $ColorAttention "Are you sure (Y/N) " -NoNewline

        If ((Read-Host).ToUpper() -ne "Y")
        {
	        Throw "[YOUR TEXT HERE] cancelled by user"
	    }
    }

    RDW-Write-Host "Begin MAIN"
    RDW-Write-Host "========================================================================"

    <#
        Code which is meant to be the body (main) of your script should be placed here.
		Make sure to set $status variable for output code (0 is success)
    #>


    RDW-Write-Host "========================================================================"
    RDW-Write-Host "End MAIN"
    return $status
}

#endregion ------------------------------------ END MAIN ------------------------------------------------

#region ---------------------------------------- Post Processing -----------------------------------------
$ExitCode = 0

If ([Environment]::GetCommandLineArgs() -Like '*-noni*')
{
    $RunInteractive = $False
} 

Else
{
    $RunInteractive = $True
}

Try 
{ 
    RDW-Start-Script
    Read-IniFile $inifile
    Read-XMLFile $xmlfile
    $ExitCode = Main $args
}

Catch
{
    $ErrorActionPreference = 'Continue'
    # Set Exitcode to 1 ensuring a direct entry into the catch statement.
    $ExitCode = 1
    RDW-Write-Host ERROR : $_.Exception.message
    RDW-Write-Host "Begin details ================================"
    $_
     
    Try
    { 
        $_ | Out-File $LogFileName -Append 
    }
    Catch
    {
        # Do nothing, we are already inside the catch 
    }

   RDW-Write-Host "End details =================================="
}

RDW-Write-Host "End script with exit code [$ExitCode]"
RDW-Stop-Script
Exit $ExitCode

#endregion---------------------------------------- End Post Processing ----------------------------------- 
