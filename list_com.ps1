# Retrieve all COM ports and their descriptions
$comPorts = Get-WmiObject Win32_SerialPort | Select-Object DeviceID, Description

# Display the COM ports and their descriptions in the console
foreach ($port in $comPorts) {
    Write-Host "$($port.DeviceID) - $($port.Description)"
}

# Wait for the user to press Enter before exiting
Read-Host -Prompt "Press Enter to exit"
