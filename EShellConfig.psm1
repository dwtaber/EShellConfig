using namespace System.Collections
using namespace System.Collections.Generic
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Security.Principal
using namespace Microsoft.Dism.Commands
using namespace Microsoft.Management.Infrastructure

#region C# types
$TypeDef = @"
using System;
using System.Runtime.InteropServices;

public static class ShellLauncherLicense
{
    const int S_OK = 0;

    public static bool IsShellLauncherLicenseEnabled()
    {
        int enabled = 0;

        if (NativeMethods.SLGetWindowsInformationDWORD("EmbeddedFeature-ShellLauncher-Enabled", out enabled) != S_OK)
        {
            enabled = 0;
        }

        return (enabled != 0);
    }

    static class NativeMethods
    {
        [DllImport("Slc.dll")]
        internal static extern int SLGetWindowsInformationDWORD([MarshalAs(UnmanagedType.LPWStr)]string valueName, out int value);
    }
}
"@
Add-Type -TypeDefinition $TypeDef
#endregion

#region Strings
$Namespace = "root/StandardCimv2/embedded"
$ClassName = "WESL_UserSetting"
$EslFeature = "Client-EmbeddedShellLauncher"
#endregion

#region Splats and selections

$SkipShouldProcess =
@{
    WhatIf  = $false
    Confirm = $false
}

$CimSplat =
@{
    Namespace   = $Namespace
    ClassName   = $ClassName
    ErrorAction = "Stop"
}

$WinFeatureSplat =
@{
    Online      = $true
    FeatureName = $EslFeature
}

$DefaultShellProps =
@(
    "Shell",
    @{ Name = "DefaultExitAction"; Expression = { [DefaultAction]$_.DefaultAction } }
)

$CustomShellProps = $DefaultShellProps + @{ Name = "Name"; Expression = { SidToName $_.Sid } }

#endregion

#region enums
enum EslMethods
{
    GetCustomShell
    SetCustomShell
    RemoveCustomShell
    GetDefaultShell
    SetDefaultShell
    IsEnabled
    SetEnabled
}

enum DefaultAction
{
    RestartShell   = 0
    RestartDevice  = 1
    ShutdownDevice = 2
    Nothing        = 3
}
#endregion

#region Parameter validation functions

# returns $true if filename or fully qualified file path is in $env:PATH
function TestEnvPathLeaf ([FileInfo]$Path)
{
    $Parent = Split-Path -Path $Path -Parent
    $Leaf = Split-Path -Path $Path -Leaf
    $EnvPathFolders = $env:Path -split ";"
    $ParentEmpty = [string]::IsNullOrWhiteSpace($Parent)
    $ParentInEnvPath = $EnvPathFolders -contains $Parent

    if ( -not ($ParentEmpty -or $ParentInEnvPath) ) { return $false }

    $TestPathResults = foreach ($Folder in $EnvPathFolders)
    {
        $JoinedPath = Join-Path -Path $Folder -ChildPath $Leaf
        Test-Path -Path $JoinedPath -PathType Leaf
    }
    if ($TestPathResults -contains $true) { return $true }
    else { return $false }
}

function ValidatePath ([FileInfo]$Path, [bool]$Force)
{
    try
    {
        switch ($true)
        {
            ($Force)
            {
                Write-Verbose -Message "-Force is set; bypassing path validation"
                break
            }
            (Test-Path -Path $Path -PathType Leaf -ErrorAction Stop)
            {
                Write-Verbose -Message "Validated fully qualified path"
                break
            }
            (TestEnvPathLeaf $Path)
            {
                Write-Verbose -Message "File exists on PATH"
                break
            }
            default
            {
                throw "Unable to validate path.  Use -Force parameter to override."
            }
        }
    }
    catch [UnauthorizedAccessException]
    {
        throw "Access to path denied.  Use -Force parameter to override."
    }
}

function ValidateName ([string]$Name)
{
    try { $null -ne (NameToSid $Name) }
    catch { throw "Could not translate user/group name to an SID." }
}
#endregion

#region SID/name functions
function SidToName([SecurityIdentifier]$Sid)
{
    return $Sid.Translate([NTAccount])
}

function NameToSid([NTAccount]$Name)
{
    return $Name.Translate([SecurityIdentifier])
}

function GetNameAndSid([string]$NameOrSid)
{
    [NTAccount]$Name = $null
    [SecurityIdentifier]$Sid  = $null

    # Works if a valid SID is provided.
    try
    {
        $Sid = $NameOrSid
        $Name = SidToName $Sid
    }
    catch
    {
        Write-Verbose -Message '$NameOrSid is not a valid SID'
    }

    # Works if valid name is provided
    try
    {
        $Sid = NameToSid $NameOrSid
        $Name = SidToName $Sid
    }
    catch
    {
        Write-Verbose -Message '$NameOrSid is not a valid name'
    }

    if ( ($null -eq $Name) -or ($null -eq $Sid) )
    { throw "Invalid name or SID" }

    return @{
        Name = $Name
        Sid = $Sid
    }
}

# Only intended for use within another function.
function SetNameAndSid ([string]$NameOrSid)
{
    $NameAndSid = GetNameAndSid $NameOrSid
    $NameAndSid.GetEnumerator() | ForEach-Object {
        Set-Variable -Name $_.Name -Value $_.Value -Scope 1 @SkipShouldProcess
    }

}

function TryParseSid ([string]$s, [ref][SecurityIdentifier]$result )
{
    try
    {
        $result.Value = [SecurityIdentifier]$s
        return $true
    }
    catch [PSInvalidCastException]
    { return $false }
}

#endregion

#region Windows feature functions
function GetEslFeatureState()
{
    return (Get-WindowsOptionalFeature @WinFeatureSplat).State
}

function IsEslFeatureEnabled()
{
    return (GetEslFeatureState) -eq [FeatureState]::Enabled
}

function IsEslFeatureLicensed()
{
    return [ShellLauncherLicense]::IsShellLauncherLicenseEnabled()
}
#endregion

#region CIM class functions
function DoesClassExist()
{
    #return $null -ne (Get-CimClass @CimSplat)
    try
    {
        $null = Get-CimClass @CimSplat
        return $true
    }
    catch
    {
        return $false
    }
}

function InvokeEslMethod([EslMethods]$MethodName, [IDictionary]$Arguments)
{
    $CimMethodSplat = $SkipShouldProcess + $CimSplat +
    @{
        MethodName = $MethodName
        Arguments  = $Arguments
    }

    Invoke-CimMethod @CimMethodSplat
}

function GetCustomShell ([string]$Sid)
{
    InvokeEslMethod ([EslMethods]::GetCustomShell) $PSBoundParameters
}

function GetCustomShellAll()
{
    Get-CimInstance @CimSplat
}

function SetCustomShell ([string]$Sid, [string]$Shell, [nullable[int]]$DefaultAction)
{
    InvokeEslMethod ([EslMethods]::SetCustomShell) $PSBoundParameters
}

function RemoveCustomShell ([string]$Sid)
{
    InvokeEslMethod ([EslMethods]::RemoveCustomShell) $PSBoundParameters
}

function GetDefaultShell()
{
    InvokeEslMethod ([EslMethods]::GetDefaultShell)
}

function SetDefaultShell([string]$Shell, [int]$DefaultAction)
{
    InvokeEslMethod ([EslMethods]::SetDefaultShell) $PSBoundParameters
}

function IsClassEnabled()
{
    (InvokeEslMethod ([EslMethods]::IsEnabled)).Enabled
}

function SetEnabled ([bool]$Enabled)
{
    InvokeEslMethod ([EslMethods]::SetEnabled) $PSBoundParameters
}
#endregion

#region ThrowIf functions
function ThrowIfUnlicensed()
{
    if (IsEslFeatureLicensed) { Write-Verbose -Message "Shell Launcher feature is licensed." }
    else { throw "Shell Launcher feature is not licensed." }
}

function ThrowIfFeatureDisabled()
{
    $FeatureState = GetEslFeatureState
    if (IsEslFeatureEnabled) { Write-Verbose -Message "Shell Launcher feature is enabled." }
    else { throw "Shell Launcher feature isn't enabled.  Current state: $FeatureState." }
}

function ThrowIfClassMissing()
{
    if (DoesClassExist) { Write-Verbose "CIM class $ClassName exists." }
    else { throw "CIM class $ClassName does not exist." }
}

function ThrowIfClassDisabled()
{
    if (IsClassEnabled) { Write-Verbose -Message "$ClassName is enabled." }
    else { throw "$ClassName is disabled." }
}

# CIM class doesn't exist until Windows feature is enabled, and is removed if the feature is disabled.
# CIM method IsEnabled reports a value different from the feature's enabled/disabled status.
# Should be possible to use CIM class's presence as a proxy for whether the feature is enabled.
function TestEslReadiness()
{
    try
    {
        ThrowIfUnlicensed
        ThrowIfClassMissing
        ThrowIfClassDisabled
    }
    catch
    {
        throw $_
    }
}

#endregion

#region Public commands

function Enable-EslClass
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param ()

    begin
    {
        ThrowIfUnlicensed
        $ClassExists = DoesClassExist
        $Target = $EslFeature
        $Operation = "Enable"
    }

    process
    {
        if ($PSCmdlet.ShouldProcess($Target,$Operation))
        {
            switch ($true)
            {
                ($ClassExists -eq $false) { Enable-WindowsOptionalFeature @WinFeatureSplat -All }
                (DoesClassExist) { SetEnabled $true }
            }
        }

        ThrowIfClassDisabled

        Write-Information -MessageData "Shell Launcher feature is enabled."
    }

    end {}
}

function Disable-EslClass
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param ()

    begin
    {
        $Target = $ClassName
        $Operation = "Disable"
        try { ThrowIfClassMissing }
        catch { Throw "Could not find CIM Class.  Check whether the Shell Launcher Windows feature is enabled." }
    }

    process
    {
        if ( $PSCmdlet.ShouldProcess($Target,$Operation) )
        {
            SetEnabled $false
            if (IsClassEnabled) { throw "Failed to disable Shell Launcher Class" }
            else { Write-Information -MessageData "Shell Launcher class is disabled." }
        }
    }

    end {}
}

function Get-EslCustomShell
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string]
        $NameOrSid
    )

    begin
    {
        TestEslReadiness
    }

    process
    {
        if ( [string]::IsNullOrWhiteSpace($NameOrSid) )
        {
            $Response = GetCustomShellAll
        }
        else
        {
            SetNameAndSid $NameOrSid
            $Response = GetCustomShell $Sid
        }
        $Response | Select-Object -Property $CustomShellProps
    }

    end {}
}

function Set-EslCustomShell
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param
    (
        [Parameter(Mandatory)]
        [string]
        $NameOrSid,

        [Parameter (Mandatory)]
        [FileInfo]
        $Path,

        [Parameter(Mandatory = $false)]
        [Alias("DefaultAction")]
        [DefaultAction]
        $DefaultExitAction,

        [Parameter()]
        [switch]
        $Force
    )

    begin
    {
        ValidatePath $Path $Force
        TestEslReadiness
        SetNameAndSid $NameOrSid
        $Target = $Name
        $Operation = "Set custom shell to $Path"
        if ($null -ne $DefaultExitAction)
        {
            $Operation += " with default exit action '$DefaultExitAction'"
        }
    }

    process
    {
        if ( $PSCmdlet.ShouldProcess($Target,$Operation) )
        {
            if ($Force)
            {
                try
                {
                    RemoveCustomShell -Sid $Sid
                }
                catch
                {
                    if
                    (
                        ($_.Exception.GetType() -eq [CimException]) -and
                        ($_.CategoryInfo.Category -eq [ErrorCategory]::ObjectNotFound)
                    )
                    {
                        Write-Verbose -Message "No previous custom shell set for $Name"
                    }
                    else { throw $_}
                }
            }
            try
            {
                SetCustomShell -Sid $Sid -Shell $Path -DefaultAction $DefaultExitAction
            }
            catch
            {
                if
                (
                    ($_.Exception.GetType() -eq [CimException]) -and
                    ($_.CategoryInfo.Category -eq [ErrorCategory]::ResourceExists)
                )
                {
                    throw "A custom shell for $Name has already been set.  Remove that configuration first or use the -Force parameter to overwrite."
                }
                else { throw $_}
            }
        }
    }

    end {}
}

function Remove-EslCustomShell
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param
    (
        [Parameter(Mandatory)]
        [string]
        $NameOrSid
    )

    begin
    {
        TestEslReadiness
        SetNameAndSid $NameOrSid
        $Target = $Name
        $Operation = "Remove custom shell configuration"
    }

    process
    {
        if ( $PSCmdlet.ShouldProcess($Target,$Operation) )
        {
            try
            {
                RemoveCustomShell $Sid
            }
            catch
        {
            if
            (
                ($_.Exception.GetType() -eq [CimException]) -and
                ($_.CategoryInfo.Category -eq [ErrorCategory]::ObjectNotFound)
            )
            {
                Write-Information -MessageData "No custom shell was set for $Name."
            }
            else
            {
                throw $_
            }
        }
        }
    }

    end {}
}

function Get-EslDefaultShell
{
    [CmdletBinding()]
    param()

    begin
    {
        TestEslReadiness
    }

    process
    {
        try { GetDefaultShell | Select-Object -Property $DefaultShellProps }
        catch
        {
            if
            (
                ($_.Exception.GetType() -eq [CimException]) -and
                ($_.CategoryInfo.Category -eq [ErrorCategory]::ObjectNotFound)
            )
            {
                throw "No default shell has been set."
            }
            else
            {
                throw $_
            }
        }
    }

    end {}
}

function Set-EslDefaultShell
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param
    (
        [Parameter (Mandatory = $true)]
        [FileInfo]
        $Path,

        [Parameter (Mandatory = $true)]
        [Alias("DefaultAction")]
        [DefaultAction]
        $DefaultExitAction,

        [Parameter()]
        [switch]
        $Force
    )

    begin
    {
        ValidatePath $Path $Force
        TestEslReadiness
        $Target = "default shell"
        $Operation = "Set path to $Path"
        if ($null -ne $DefaultExitAction)
        {
            $Operation += " with default exit action '$DefaultExitAction'"
        }
    }

    process
    {
        if ( $PSCmdlet.ShouldProcess($Target,$Operation) )
        {
            SetDefaultShell $Path $DefaultExitAction
        }
    }

    end {}
}

#endregion
