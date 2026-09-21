# Define Hostname and Username
$Hostname = $env:COMPUTERNAME
$Username = $env:USERNAME
$OutputFile = "$HOME\Desktop\${Hostname}_${Username}.txt"

# Gather System Information
$Serial = (Get-CimInstance Win32_BIOS).SerialNumber
$Model = (Get-CimInstance Win32_ComputerSystem).Model
$OS = (Get-CimInstance Win32_OperatingSystem).Caption
$CPU = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name

# Calculate RAM in GB
$RAM_Bytes = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
$RAM_GB = [math]::Round($RAM_Bytes / 1GB, 2)

# Get Physical Storage Details
$Storage = (Get-PhysicalDisk | Where-Object MediaType -match 'SSD|HDD' | ForEach-Object { "$($_.FriendlyName) - $([math]::Round($_.Size / 1GB, 2)) GB" }) -join "; "

# Format the Output
$Output = @"
IT Inventory Details
-----------------------
Hostname     : $Hostname
Username     : $Username
Serial Number: $Serial
Model Number : $Model
OS Version   : $OS
CPU          : $CPU
Memory (RAM) : $RAM_GB GB
Storage      : $Storage
"@

# Save to File
$Output | Out-File -FilePath $OutputFile -Encoding utf8
Write-Host "Success! Your inventory file has been saved to your Desktop as: ${Hostname}_${Username}.txt" -ForegroundColor Green