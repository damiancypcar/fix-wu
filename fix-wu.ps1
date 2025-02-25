# ----------------------------------------------------------
# Author:          damiancypcar
# Modified:        2025-02-25
# Version:         1.0
# Desc:            Fixes Windows Update stuck problem
# ----------------------------------------------------------
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]

# Fixes Windows Update stuck problem
Param (
    [switch]$help = $false
)

# All error stop the process
$ErrorActionPreference = "Stop"

# ROOT CHECK
$ErrorActionPreference = "Stop";
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be executed as Administrator.";
    exit 1;
}

function Show-Help {
    Write-Output "`nFixes Windows Update stuck problem"
    Write-Output "`nUsage:"

    Write-Output "`n1. Run PowerShell as Administrator"
    Write-Output "`n2. Set the execution policy"
    Write-Output "    >  Set-ExecutionPolicy Bypass -Scope Process -Force"
    Write-Output "`n3. Run script directly from the web"
    Write-Output "    >  irm 'https://raw.githubusercontent.com/damiancypcar/fix-wu/refs/heads/main/fix-wu.ps1' | iex"
    break
}

function Stop-WUServices {
    # Stop Windows Update Services
    Stop-Service -Name wuauserv -Force
    Stop-Service -Name cryptSvc -Force
    Stop-Service -Name bits -Force
    Stop-Service -Name msiserver -Force
}

function Remove-WUCache {
    # Delete Windows Update Cache
    Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force
    Remove-Item -Path "C:\Windows\System32\catroot2" -Recurse -Force
}

function Start-WUServices {
    # Start Windows Update Services
    Start-Service -Name wuauserv
    Start-Service -Name cryptSvc
    Start-Service -Name bits
    Start-Service -Name msiserver    
}


function Main {
    Write-Output "`nFixes Windows Update stuck problem"
    $confirmation = Read-Host "`nAre you sure you want to proceed (y/N)?"
    if ($confirmation -eq 'y') {
        Stop-WUServices
        Remove-WUCache
        Start-WUServices
        Write-Output "Done"
        Write-Output "Restart the Computer"
    } else {
        Write-Output "Exiting"
        break
    }
}

if ($help) {
    Show-Help
}else {
    Main
}
