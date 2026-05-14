<#
.SYNOPSIS
    Reports the presence of VMware Tools artifacts.

    Updates: 
    5/14/2026 - corrected some path errors
    
#>

function Get-VMwareToolsInstallerID {
    # HKCR is a merged hive; searching it once is usually sufficient
    $path = "Registry::HKEY_CLASSES_ROOT\Installer\Products"
    Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.GetValue('ProductName') -eq 'VMware Tools') {
            return @{
                reg_id = $_.PSChildName;
                msi_id = [Regex]::Match($_.GetValue('ProductIcon'), '(?<={)(.*?)(?=})').Value
            }
        }
    }
}

$vmware_ids = Get-VMwareToolsInstallerID

# We assign the loop output directly to $Report to avoid += performance hits
$Report = & {
    # 1. Registry Targets (Condensed list)
    $reg_paths = @(
        "HKLM:\SOFTWARE\Classes\Installer\Products\",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" # Added WOW6432Node
    )

    if ($vmware_ids) {
        foreach ($path in $reg_paths) {
            $fullPath = Join-Path $path $vmware_ids.reg_id
            [PSCustomObject]@{Type="Registry"; Target=$fullPath; Found=(Test-Path $fullPath)}
        }
        
        $msiPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{$($vmware_ids.msi_id)}"
        [PSCustomObject]@{Type="Registry"; Target=$msiPath; Found=(Test-Path $msiPath)}
    }

    # 2. Legacy Keys
    if ([Environment]::OSVersion.Version.Major -lt 10) {
        $legacy = @(
            "HKCR:\CLSID\{D86ADE52-C4D9-4B98-AA0D-9B0C7F1EBBC8}",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9709436B-5A41-4946-8BE7-2AA433CAF108}"
        )
        foreach ($k in $legacy) { [PSCustomObject]@{Type="Registry (Legacy)"; Target=$k; Found=(Test-Path $k)} }
    }

    # 3. Directories
    $dirs = @("C:\Program Files\VMware", "C:\Program Files\Common Files\VMware")
    foreach ($d in $dirs) { [PSCustomObject]@{Type="Directory"; Target=$d; Found=(Test-Path $d)} }

    # 4. Services
    Get-Service -Name "VMware*", "GISvc" -ErrorAction SilentlyContinue | ForEach-Object {
        [PSCustomObject]@{Type="Service"; Target=$_.Name; Found=$true}
    }
}

# --- Output ---
Write-Host "`n--- VMware Tools Presence Report ---" -ForegroundColor Cyan
$Report | Format-Table -AutoSize

$foundCount = ($Report | Where-Object { $_.Found }).Count
if ($foundCount -gt 0) {
    Write-Host "STATUS: DIRTY ($foundCount artifacts found)" -ForegroundColor Red
} else {
    Write-Host "STATUS: CLEAN" -ForegroundColor Green
}
