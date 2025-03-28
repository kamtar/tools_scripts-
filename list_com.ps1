# Retrieve all COM ports and their descriptions using Win32_PnPEntity
$comPorts = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Name -match "COM\d+" }

# Display the COM ports, descriptions, and driver provider in the console
foreach ($port in $comPorts) {
    $comPortName = [regex]::Match($port.Name, "COM\d+").Value  # Extract COM port (e.g., COM6)
    $cleanName = $port.Name -replace "\s*\($comPortName\)", "" # Remove (COMx) from the name
    $driverProvider = $port.Manufacturer   # Get the driver provider
    
    if (-not $driverProvider) { $driverProvider = "Unknown" }  # Fallback if provider is missing
    
    Write-Host "$comPortName - $cleanName ($driverProvider)"
}

# Wait for the user to press Enter before exiting
Read-Host -Prompt "Press Enter to exit"
