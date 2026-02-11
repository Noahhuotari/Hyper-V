#Get all Virtual Switches and their associated physical adapters
$Switches = Get-VMSwitch

foreach ($Switch in $Switches) {
    Write-Host "--- Switch Name: $($Switch.Name) ---" -ForegroundColor Cyan
    Write-Host "Switch Type: $($Switch.SwitchType)"
    Write-Host "Embedded Teaming Enabled: $($Switch.EmbeddedTeamingEnabled)"
    
    # Get the physical adapters bound to this switch
    $NetAdapters = Get-VMSwitchTeam -SwitchName $Switch.Name
    
    if ($NetAdapters) {
        Write-Host "Physical Adapters in Team:" -ForegroundColor Yellow
        $NetAdapters.NetAdapterInterfaceDescription | ForEach-Object {
            $AdapterDetail = Get-NetAdapter -InterfaceDescription $_
            Write-Host " > [$($AdapterDetail.Status)] $_ (Name: $($AdapterDetail.Name))"
        }
    } else {
        # Fallback for standard (non-SET) external switches
        $Extension = Get-VMSwitchExtension -SwitchName $Switch.Name | Where-Object { $_.Id -eq "Microsoft.VMSwitch.External" }
        Write-Host "Physical Adapter: $($Switch.NetAdapterInterfaceDescription)" -ForegroundColor Yellow
    }
    Write-Host ""
}
