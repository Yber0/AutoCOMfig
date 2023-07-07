function ConfigureCOMObject (
            $HardwareID,      # To search for COM objects a given model
            $NewFriendlyName, # What the object will be displayed as after configured
            $NewPort,         # COM port we want a device to be configured to use
            $Configuration    # Bits por segundo, Bits de datos, Paridad, Bits de parada y Control de flujo
    )
    {
   
    $Devices = Get-PnpDevice -PresentOnly | Where-Object { $_.HardwareID -match $HardwareID }
    # Foreach connected COM
    foreach ($Device in $Devices)
    {
        # Change basic Configuration
        $COMPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports"
        Set-ItemProperty -Path $COMPath -Name "$($NewPort):" -Value $XMLEntry.configuration

        # Get currently used COM ports (Plugged PnPDevices with the PortName parameter)
        $COMObjects = Get-PnpDevice -PresentOnly | Where-Object { (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Enum\$($_.DeviceID)\Device Parameters" -Name "PortName" -ErrorAction SilentlyContinue) -ne $null }
       
        # COM ports being used  
        $COMBusy = [System.Collections.ArrayList]@()
        # DeviceIDs using ports      
        $COMDevice = [System.Collections.ArrayList]@()
        # Fill arrays with current COM objects
        foreach ($COMObject in $COMObjects) {
            $COMBusy.Add((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Enum\$($COMObject.DeviceID)\Device Parameters" -Name "PortName").PortName) | Out-Null
            $COMDevice.Add($COMObject.DeviceID) | Out-Null
        }
       
        # Other device is using the port we're trying to set
        $taken = $false
        # This device is using the port we're trying to set
        $using = $false
        # Check if COM port is being used
        for ($i = 0;$i -lt $COMBusy.Count;$i++) {
            if ($COMBusy[$i] -contains $NewPort) {
                if ($COMDevice[$i] -eq $Device.DeviceID) { $using = $true }
                else { $taken = $true }
            }
        }

        # Check error statuses
        if ($taken -and $using) {
            Write-Host -ForegroundColor Red "Error:" $Device.Name "and other devices are using the port $($NewPort)"
            continue
        }
        if ($taken -and -not $using) {
            Write-Host -ForegroundColor Red "Error:" $Device.Name "-" $($NewPort) "beign used by other device"
            continue
        }
        if ($using -and -not $taken) {
            Write-Host $Device.Name "alredy using the port $($NewPort)"
            continue
        }

        # Set new port and update Friendly Name
        $PortKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $Device.DeviceID + "\Device Parameters"
        $DeviceKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $Device.DeviceID
        $FriendlyName = $NewFriendlyName + " (" + $NewPort + ")"
        New-ItemProperty -Path $PortKey -Name "PortName" -PropertyType String -Value $NewPort -Force | Out-Null
        New-ItemProperty -Path $DeviceKey -Name "FriendlyName" -PropertyType String -Value $FriendlyName -Force | Out-Null

        #Release Previous Com Port from ComDB
        $Byte = ($CurrentPort - ($CurrentPort % 8))/8
        $Bit = 8 - ($CurrentPort % 8)
        if ($Bit -eq 8) {
            $Bit = 0
            $Byte = $Byte - 1
        }

        # Update Port Change in COMDB
        $ComDB = get-itemproperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\COM Name Arbiter" -Name ComDB
        $ComBinaryArray = ([convert]::ToString($ComDB.ComDB[$Byte],2)).ToCharArray()

        while ($ComBinaryArray.Length -ne 8) { $ComBinaryArray = ,"0" + $ComBinaryArray }

        $ComBinaryArray[$Bit] = "0"
        $ComBinary = [string]::Join("",$ComBinaryArray)
        $ComDB.ComDB[$Byte] = [convert]::ToInt32($ComBinary,2)

        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\COM Name Arbiter" -Name ComDB -Value ([byte[]]$ComDB.ComDB)

        # Disable/Enable device (refresh)
        Get-PnpDevice -PresentOnly -FriendlyName "$($FriendlyName)*" | Disable-PnpDevice -Confirm:$false
        Start-Sleep -Seconds 5
        Get-PnpDevice -PresentOnly -FriendlyName "$($FriendlyName)*" | Enable-PnpDevice -Confirm:$false
        Write-Host $Device.Name " has been configured. New Name: " $FriendlyName
    }
}
function CallConfiguration {
    Write-Host `r`n(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    # Configure COM devices listed inside the XML
    foreach ($XMLEntry in $XMLDevices) { ConfigureCOMObject $XMLEntry.hardwareid $XMLEntry.name $XMLEntry.newport $XMLEntry.configuration }
}

# Read XML config values
$XMLDevices = Select-Xml -Path .\AutoCOMfig.xml -XPath /Devices/Device | ForEach-Object { $_.Node }

# Configure currently connected COM devices once before waiting for event
CallConfiguration