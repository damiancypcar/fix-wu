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

    Write-Output "`nStopping Windows Update services"

    Write-Output "  > Stopping Windows Update Service (wuauserv)..."
    Stop-Service -Name wuauserv -Force

    Write-Output "  > Stopping Cryptographic Services (cryptSvc)..."
    Stop-Service -Name cryptSvc -Force

    Write-Output "  > Stopping Background Intelligent Transfer Service (BITS)..."
    Stop-Service -Name bits -Force

    Write-Output "  > Stopping Windows Installer Service (msiserver)..."
    Stop-Service -Name msiserver -Force
}

function Remove-WUCache {
    # Delete Windows Update Cache
    Write-Output "`nRemoving Windows Update cache"

    Write-Output "  > Removing SoftwareDistribution folder..."
    Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force

    Write-Output "  > Removing catroot2 folder..."
    Remove-Item -Path "C:\Windows\System32\catroot2" -Recurse -Force
}

function Start-WUServices {
    # Start Windows Update Services
    Write-Output "`nStarting Windows Update services"

    Write-Output "  > Starting Windows Update Service (wuauserv)..."
    Start-Service -Name wuauserv

    Write-Output "  > Starting Cryptographic Services (cryptSvc)..."
    Start-Service -Name cryptSvc

    Write-Output "  > Starting Background Intelligent Transfer Service (BITS)..."
    Start-Service -Name bits

    Write-Output "  > Starting Windows Installer Service (msiserver)..."
    Start-Service -Name msiserver
}


function Main {
    Write-Output "`nFixes Windows Update stuck problem"
    $confirmation = Read-Host "`nAre you sure you want to proceed (y/N)?"
    if ($confirmation -eq 'y') {
        Stop-WUServices
        Remove-WUCache
        Start-WUServices
        Write-Output "`nDone"
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
