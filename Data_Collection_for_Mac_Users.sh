# Define Hostname and Username
HOST_NAME=$(scutil --get LocalHostName)
USER_NAME=$(whoami)
OUTPUT_FILE="$HOME/Desktop/${HOST_NAME}_${USER_NAME}.txt"

# Gather System Information
SERIAL=$(system_profiler SPHardwareDataType | awk -F': ' '/Serial Number/{print $2}')
MODEL=$(system_profiler SPHardwareDataType | awk -F': ' '/Model Identifier/{print $2}')
OS="$(sw_vers -productName) $(sw_vers -productVersion)"
CPU=$(sysctl -n machdep.cpu.brand_string)
RAM=$(system_profiler SPHardwareDataType | awk -F': ' '/Memory/{print $2}')
STORAGE=$(diskutil info / | grep "Disk Size" | cut -d: -f2 | sed 's/^ *//')

# Format and Save to File
cat <<EOF > "$OUTPUT_FILE"
IT Inventory Details
-----------------------
Hostname     : $HOST_NAME
Username     : $USER_NAME
Serial Number: $SERIAL
Model Number : $MODEL
OS Version   : $OS
CPU          : $CPU
Memory (RAM) : $RAM
Storage      : $STORAGE
EOF

echo "Success! Your inventory file has been saved to your Desktop as: ${HOST_NAME}_${USER_NAME}.txt"