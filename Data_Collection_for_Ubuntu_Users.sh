#!/bin/bash

# Define Hostname and Username
HOST_NAME=$(hostname)
USER_NAME=$(whoami)
OUTPUT_FILE="$HOME/Desktop/${HOST_NAME}_${USER_NAME}.txt"

# Ensure Desktop directory exists
mkdir -p "$HOME/Desktop"

# Gather System Information
SERIAL=$(sudo dmidecode -s system-serial-number 2>/dev/null || echo "N/A")
MODEL=$(sudo dmidecode -s system-product-name 2>/dev/null || echo "N/A")
OS=$(lsb_release -ds 2>/dev/null || cat /etc/os-release | grep PRETTY_NAME | cut -d'"' -f2)
CPU=$(lscpu | grep "Model name" | cut -d: -f2 | sed 's/^ *//')
RAM=$(free -h | awk 'NR==2{print $2}')
STORAGE=$(df -h / | awk 'NR==2{print $2}')

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
