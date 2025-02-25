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

function Show-Help {
    Write-Output "`nFixes Windows Update stuck problem"
    Write-Output "`nUsage:"

    Write-Output "`n1. Run PowerShell as Administrator"
    Write-Output "`n2. Set the execution policy"
    Write-Output "    >  Set-ExecutionPolicy Bypass -Scope Process -Force"
    Write-Output "`n3. Run script directly from the web"
    Write-Output "    >  irm http://webpage.com/script.ps1 | iex"
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
    Stop-WUServices
    Remove-WUCache
    Start-WUServices
    Write-Output "Restart the Computer"
}

if ($help) {
    Show-Help
}else {
    Main
}
