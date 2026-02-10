<#
.SYNOPSIS
    Reports the presence of VMware Tools artifacts (Registry, Files, Services).
    Based on the manual removal logic for Windows 2008-2019.
#>

function Get-VMwareToolsInstallerID {
    foreach ($item in $(Get-ChildItem Registry::HKEY_CLASSES_ROOT\Installer\Products -ErrorAction SilentlyContinue)) {
        if ($item.GetValue('ProductName') -eq 'VMware Tools') {
            return @{
                reg_id = $item.PSChildName;
                msi_id = [Regex]::Match($item.GetValue('ProductIcon'), '(?<={)(.*?)(?=})') | Select-Object -ExpandProperty Value
            }
        }
    }
}

$vmware_tools_ids = Get-VMwareToolsInstallerID
$VMware_Tools_Directory = "C:\Program Files\VMware"
$VMware_Common_Directory = "C:\Program Files\Common Files\VMware"
$Report = @()

# 1. Check Registry Targets
$reg_paths = @(
    "HKCR:\Installer\Features\",
    "HKCR:\Installer\Products\",
    "HKLM:\SOFTWARE\Classes\Installer\Features\",
    "HKLM:\SOFTWARE\Classes\Installer\Products\",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\"
)

if ($vmware_tools_ids) {
    foreach ($path in $reg_paths) {
        $fullPath = Join-Path $path $vmware_tools_ids.reg_id
        $exists = Test-Path $fullPath
        $Report += [PSCustomObject]@{Type="Registry"; Target=$fullPath; Found=$exists}
    }
    
    $msiPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{$($vmware_tools_ids.msi_id)}"
    $Report += [PSCustomObject]@{Type="Registry"; Target=$msiPath; Found=(Test-Path $msiPath)}
}

# 2. Check Legacy/OS Specific Keys
if ([Environment]::OSVersion.Version.Major -lt 10) {
    $legacyKeys = @(
        "HKCR:\CLSID\{D86ADE52-C4D9-4B98-AA0D-9B0C7F1EBBC8}",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9709436B-5A41-4946-8BE7-2AA433CAF108}",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}"
    )
    foreach ($key in $legacyKeys) {
        $Report += [PSCustomObject]@{Type="Registry (Legacy)"; Target=$key; Found=(Test-Path $key)}
    }
}

# 3. Check General VMware Registry
$vmwareInc = "HKLM:\SOFTWARE\VMware, Inc."
$Report += [PSCustomObject]@{Type="Registry"; Target=$vmwareInc; Found=(Test-Path $vmwareInc)}

# 4. Check Filesystem
$Report += [PSCustomObject]@{Type="Directory"; Target=$VMware_Tools_Directory; Found=(Test-Path $VMware_Tools_Directory)}
$Report += [PSCustomObject]@{Type="Directory"; Target=$VMware_Common_Directory; Found=(Test-Path $VMware_Common_Directory)}

# 5. Check Services
$serviceNames = @("VMware*", "GISvc")
foreach ($n in $serviceNames) {
    $foundServices = Get-Service -Name $n -ErrorAction SilentlyContinue
    if ($foundServices) {
        foreach ($s in $foundServices) {
            $Report += [PSCustomObject]@{Type="Service"; Target=$s.Name; Found=$true}
        }
    }
}

# --- Output Report ---
Write-Host "`n--- VMware Tools Presence Report ---" -ForegroundColor Cyan
$Report | Format-Table -AutoSize

$dirtyItems = ($Report | Where-Object { $_.Found -eq $true }).Count

if ($dirtyItems -gt 0) {
    Write-Host "STATUS: DIRTY ($dirtyItems artifacts found)" -ForegroundColor Red
} else {
    Write-Host "STATUS: CLEAN (No artifacts found)" -ForegroundColor Green
}
