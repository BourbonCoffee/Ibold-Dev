<#
.SYNOPSIS
    Creates a new Hyper-V virtual machine with configurable settings and prepares it for remote management.

.DESCRIPTION
    This script automates the process of creating a new Hyper-V lab environment including:
    - Enabling WinRM for remote PowerShell access
    - Installing Hyper-V if not already installed
    - Creating a virtual network with NAT
    - Creating a Generation 2 VM with configurable settings
    - Sets up the VM for remote DSC configuration

.PARAMETER VMName
    The name of the virtual machine to create.

.PARAMETER MemoryStartupGB
    The amount of startup memory in GB for the virtual machine. Defaults to 4GB.

.PARAMETER VHDSizeGB
    The size of the virtual hard disk in GB. Defaults to 60GB.

.PARAMETER VHDPath
    The path where the virtual hard disk will be stored. If not specified, defaults to C:\ProgramData\Microsoft\Windows\Virtual Hard Disks\VMName.vhdx.

.PARAMETER ISOPath
    The path to the ISO file to be mounted to the virtual machine.

.PARAMETER SwitchName
    The name of the virtual switch to connect the VM to. Default is "Lab".

.PARAMETER VMPath
    The path where the VM configuration will be stored. Defaults to C:\ProgramData\Microsoft\Windows\Hyper-V.

.PARAMETER VMIP
    The IP address to assign to the virtual machine. Defaults to 172.16.0.10.

.PARAMETER Subnet
    The subnet for the virtual network. Defaults to 172.16.0.0/16.

.PARAMETER GatewayIP
    The IP address of the gateway for the virtual network. Defaults to 172.16.0.1.

.PARAMETER CPUCount
    The number of virtual processors to assign to the VM. Defaults to 2.

.PARAMETER EnableDynamicMemory
    Whether to enable dynamic memory for the VM. Defaults to $true.

.PARAMETER MinimumMemoryMB
    The minimum amount of memory in MB when dynamic memory is enabled. Defaults to 512MB.

.PARAMETER MaximumMemoryGB
    The maximum amount of memory in GB when dynamic memory is enabled. Defaults to 8GB.

.EXAMPLE
    .\New-HypervLabEnvironment.ps1 -VMName "DC01" -MemoryStartupGB 4 -VHDSizeGB 80 -ISOPath "C:\ISO\Windows_Server_2019.iso"

.EXAMPLE
    .\New-HypervLabEnvironment.ps1 -VMName "SQL01" -MemoryStartupGB 8 -VHDSizeGB 120 -CPUCount 4 -VMIP "172.16.0.20"

.NOTES
    Author: Chris Ibold
    Company: Comet Consulting Group
    Version: 1.0
    Date: 2025-04-24
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$VMName,

    [Parameter(Mandatory = $false)]
    [int]$MemoryStartupGB = 4,

    [Parameter(Mandatory = $false)]
    [int]$VHDSizeGB = 60,

    [Parameter(Mandatory = $false)]
    [string]$VHDPath = "C:\Hyper-V\VHDs\$VMName.vhdx",

    [Parameter(Mandatory = $true)]
    [string]$ISOPath,

    [Parameter(Mandatory = $false)]
    [string]$SwitchName = "Lab",

    [Parameter(Mandatory = $false)]
    [string]$VMPath = "C:\Hyper-V\VMs",

    [Parameter(Mandatory = $false)]
    [string]$VMIP = "172.16.0.20",

    [Parameter(Mandatory = $false)]
    [string]$Subnet = "172.16.0.0/16",

    [Parameter(Mandatory = $false)]
    [string]$GatewayIP = "172.16.0.1",

    [Parameter(Mandatory = $false)]
    [int]$CPUCount = 2,

    [Parameter(Mandatory = $false)]
    [bool]$EnableDynamicMemory = $true,

    [Parameter(Mandatory = $false)]
    [int]$MinimumMemoryMB = 512,

    [Parameter(Mandatory = $false)]
    [int]$MaximumMemoryGB = 8
)

#Region Functions

function Test-AdminPrivileges {
    <#
    .SYNOPSIS
        Checks if the current PowerShell session is running with administrative privileges.
    
    .DESCRIPTION
        This function verifies if the current PowerShell session has the necessary administrative 
        privileges required for operations like installing Windows features and managing Hyper-V.
    
    .EXAMPLE
        if (-not (Test-AdminPrivileges)) {
            Write-Error "This script requires administrative privileges."
            exit
        }
    
    .OUTPUTS
        System.Boolean - Returns $true if running with admin privileges, $false otherwise.
    #>
    
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Initialize-WinRM {
    <#
    .SYNOPSIS
        Configures WinRM service for remote PowerShell connectivity.
    
    .DESCRIPTION
        This function starts the WinRM service if it's not already running and adds the specified
        IP address to the TrustedHosts list to enable remote PowerShell connections.
    
    .PARAMETER IPAddress
        The IP address to add to the TrustedHosts list.
    
    .EXAMPLE
        Initialize-WinRM -IPAddress "172.16.0.10"
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$IPAddress
    )

    try {
        # Start WinRM service if not already running
        $winRMService = Get-Service -Name WinRM
        if ($winRMService.Status -ne 'Running') {
            Write-Verbose "Starting WinRM service..."
            Start-Service -Name WinRM
            Write-Verbose "WinRM service started successfully."
        } else {
            Write-Verbose "WinRM service is already running."
        }

        # Check current TrustedHosts
        $currentTrustedHosts = Get-Item -Path WSMan:\localhost\Client\TrustedHosts | Select-Object -ExpandProperty Value
        Write-Verbose "Current TrustedHosts: $currentTrustedHosts"

        # Add the IP to TrustedHosts if not already there
        if ($currentTrustedHosts -notlike "*$IPAddress*") {
            Write-Verbose "Adding $IPAddress to TrustedHosts..."
            
            # If there are existing entries, append the new one
            if ($currentTrustedHosts) {
                $newTrustedHosts = "$currentTrustedHosts,$IPAddress"
            } else {
                $newTrustedHosts = $IPAddress
            }
            
            Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value $newTrustedHosts -Force
            Write-Verbose "$IPAddress added to TrustedHosts successfully."
        } else {
            Write-Verbose "$IPAddress is already in TrustedHosts."
        }
        
        Write-Host "WinRM configured successfully!" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to initialize WinRM: $_"
        return $false
    }
}

function Install-HyperVFeature {
    <#
    .SYNOPSIS
        Ensures the Hyper-V feature is installed on the host.
    
    .DESCRIPTION
        This function checks if Hyper-V is installed, and if not, installs it.
        It requires a restart of the computer to complete the installation.
    
    .EXAMPLE
        Install-HyperVFeature
    
    .OUTPUTS
        System.Boolean - Returns $true if Hyper-V is already installed or was successfully installed,
                        $false if installation failed.
    #>
    
    [CmdletBinding()]
    param()

    try {
        # Check if Hyper-V is already installed
        $hyperVFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All
        
        if ($hyperVFeature.State -eq "Enabled") {
            Write-Verbose "Hyper-V is already installed."
            return $true
        }
        
        # Install Hyper-V
        Write-Verbose "Installing Hyper-V. This may take a while..."
        $installResult = Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All -NoRestart
        
        if ($installResult.RestartNeeded) {
            Write-Warning "A restart is required to complete the Hyper-V installation."
            $restartChoice = Read-Host "Do you want to restart now? (Y/N)"
            
            if ($restartChoice -eq 'Y' -or $restartChoice -eq 'y') {
                Restart-Computer -Force
            } else {
                Write-Warning "Please restart your computer to complete the Hyper-V installation before continuing with this script."
                return $false
            }
        }
        
        Write-Host "Hyper-V installed successfully!" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to install Hyper-V: $_"
        return $false
    }
}

function New-VirtualNetwork {
    <#
    .SYNOPSIS
        Creates a new virtual network for Hyper-V VMs.
    
    .DESCRIPTION
        This function creates a new internal virtual switch for Hyper-V and configures NAT networking
        for the VMs to have internet access through the host.
    
    .PARAMETER SwitchName
        The name to assign to the new virtual switch.
    
    .PARAMETER SubnetPrefix
        The subnet in CIDR notation (e.g., "172.16.0.0/16").
    
    .PARAMETER GatewayIP
        The IP address to assign to the virtual switch interface on the host.
    
    .EXAMPLE
        New-VirtualNetwork -SwitchName "Lab" -SubnetPrefix "172.16.0.0/16" -GatewayIP "172.16.0.1"
    
    .OUTPUTS
        System.Boolean - Returns $true if the virtual network was created successfully, $false otherwise.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SwitchName,
        
        [Parameter(Mandatory = $true)]
        [string]$SubnetPrefix,
        
        [Parameter(Mandatory = $true)]
        [string]$GatewayIP
    )

    try {
        # Check if switch already exists
        $existingSwitch = Get-VMSwitch -Name $SwitchName -ErrorAction SilentlyContinue
        
        if ($existingSwitch) {
            Write-Verbose "Virtual switch '$SwitchName' already exists."
        } else {
            # Create new internal virtual switch
            Write-Verbose "Creating new internal virtual switch '$SwitchName'..."
            New-VMSwitch -SwitchName $SwitchName -SwitchType Internal -ErrorAction Stop
            Write-Verbose "Virtual switch created successfully."
        }
        
        # Get the network adapter for the switch
        $netAdapter = Get-NetAdapter | Where-Object { $_.Name -like "*$SwitchName*" }
        
        if (-not $netAdapter) {
            throw "Could not find network adapter for the virtual switch."
        }
        
        # Check if IP is already configured
        $existingIP = Get-NetIPAddress -InterfaceIndex $netAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
        
        if ($existingIP -and $existingIP.IPAddress -eq $GatewayIP) {
            Write-Verbose "IP address $GatewayIP is already configured on the adapter."
        } else {
            # Configure IP address on the adapter
            Write-Verbose "Configuring IP address $GatewayIP on the adapter..."
            
            # Remove existing IP if any
            if ($existingIP) {
                Remove-NetIPAddress -InterfaceIndex $netAdapter.ifIndex -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
            }
            
            # Add new IP
            $prefixLength = ($SubnetPrefix -split '/')[1]
            New-NetIPAddress -IPAddress $GatewayIP -PrefixLength $prefixLength -InterfaceIndex $netAdapter.ifIndex -ErrorAction Stop
            Write-Verbose "IP address configured successfully."
        }
        
        # Check if NAT is already configured
        $existingNat = Get-NetNat -Name "HyperVNAT" -ErrorAction SilentlyContinue
        
        if ($existingNat) {
            Write-Verbose "NAT is already configured."
            
            # Check if NAT is configured for the correct subnet
            if ($existingNat.InternalIPInterfaceAddressPrefix -ne $SubnetPrefix) {
                Write-Verbose "Updating NAT configuration..."
                Remove-NetNat -Name "HyperVNAT" -Confirm:$false
                New-NetNat -Name "HyperVNAT" -InternalIPInterfaceAddressPrefix $SubnetPrefix -ErrorAction Stop
                Write-Verbose "NAT updated to use subnet $SubnetPrefix."
            }
        } else {
            # Configure NAT
            Write-Verbose "Configuring NAT for subnet $SubnetPrefix..."
            New-NetNat -Name "HyperVNAT" -InternalIPInterfaceAddressPrefix $SubnetPrefix -ErrorAction Stop
            Write-Verbose "NAT configured successfully."
        }
        
        Write-Host "Virtual network setup complete!" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to create virtual network: $_"
        return $false
    }
}

function New-HyperVLabVM {
    <#
    .SYNOPSIS
        Creates a new Hyper-V virtual machine with specified configuration.
    
    .DESCRIPTION
        This function creates a new Generation 2 Hyper-V virtual machine with the specified
        settings, including memory, CPU, storage, and networking configuration.
    
    .PARAMETER VMName
        The name of the virtual machine to create.
    
    .PARAMETER MemoryStartupMB
        The amount of startup memory in MB for the virtual machine.
    
    .PARAMETER VHDPath
        The path where the virtual hard disk will be stored.
    
    .PARAMETER VHDSizeBytes
        The size of the virtual hard disk in bytes.
    
    .PARAMETER SwitchName
        The name of the virtual switch to connect the VM to.
    
    .PARAMETER VMPath
        The path where the VM configuration will be stored.
    
    .PARAMETER ISOPath
        The path to the ISO file to be mounted to the virtual machine.
    
    .PARAMETER ProcessorCount
        The number of virtual processors to assign to the VM.
    
    .PARAMETER EnableDynamicMemory
        Whether to enable dynamic memory for the VM.
    
    .PARAMETER MinimumMemoryMB
        The minimum amount of memory in MB when dynamic memory is enabled.
    
    .PARAMETER MaximumMemoryMB
        The maximum amount of memory in MB when dynamic memory is enabled.
    
    .EXAMPLE
        New-HyperVLabVM -VMName "TestVM" -MemoryStartupMB 4096 -VHDPath "C:\VMs\TestVM.vhdx" -VHDSizeBytes 64GB -SwitchName "Lab" -VMPath "C:\VMs" -ISOPath "C:\ISO\Windows.iso" -ProcessorCount 2 -EnableDynamicMemory $true -MinimumMemoryMB 512 -MaximumMemoryMB 8192
    
    .OUTPUTS
        System.Boolean - Returns $true if the VM was created successfully, $false otherwise.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMName,
        
        [Parameter(Mandatory = $true)]
        [int]$MemoryStartupMB,
        
        [Parameter(Mandatory = $true)]
        [string]$VHDPath,
        
        [Parameter(Mandatory = $true)]
        [int64]$VHDSizeBytes,
        
        [Parameter(Mandatory = $true)]
        [string]$SwitchName,
        
        [Parameter(Mandatory = $true)]
        [string]$VMPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ISOPath,
        
        [Parameter(Mandatory = $false)]
        [int]$ProcessorCount = 2,
        
        [Parameter(Mandatory = $false)]
        [bool]$EnableDynamicMemory = $true,
        
        [Parameter(Mandatory = $false)]
        [int]$MinimumMemoryMB = 512,
        
        [Parameter(Mandatory = $false)]
        [int]$MaximumMemoryMB = 8192
    )

    try {
        # Check if VM already exists
        $existingVM = Get-VM -Name $VMName -ErrorAction SilentlyContinue
        
        if ($existingVM) {
            Write-Warning "A VM with the name '$VMName' already exists. Please choose a different name or remove the existing VM."
            return $false
        }
        
        # Ensure the VM directory exists
        if (-not (Test-Path -Path $VMPath)) {
            Write-Verbose "Creating VM directory at $VMPath..."
            New-Item -ItemType Directory -Path $VMPath -Force | Out-Null
            Write-Verbose "VM directory created."
        }
        
        # Check if VHD already exists
        if (Test-Path -Path $VHDPath) {
            Write-Warning "A virtual hard disk already exists at '$VHDPath'. Using a new VHD may result in data loss."
            $overwriteVHD = Read-Host "Do you want to overwrite the existing VHD? (Y/N)"
            
            if ($overwriteVHD -eq 'Y' -or $overwriteVHD -eq 'y') {
                Write-Verbose "Removing existing VHD..."
                Remove-Item -Path $VHDPath -Force
                Write-Verbose "Existing VHD removed."
            } else {
                Write-Warning "VM creation aborted to preserve existing VHD."
                return $false
            }
        }
        
        # Create the VM
        Write-Verbose "Creating virtual machine '$VMName'..."
        New-VM -Name $VMName -Generation 2 -MemoryStartupBytes $MemoryStartupMB -MB -SwitchName $SwitchName -Path $VMPath -ErrorAction Stop
        Write-Verbose "Virtual machine created successfully."
        
        # Configure CPU
        Write-Verbose "Setting processor count to $ProcessorCount..."
        Set-VM -Name $VMName -ProcessorCount $ProcessorCount -ErrorAction Stop
        Write-Verbose "Processor count set successfully."
        
        # Configure memory
        if ($EnableDynamicMemory) {
            Write-Verbose "Enabling dynamic memory with range $MinimumMemoryMB MB to $MaximumMemoryMB MB..."
            Set-VM -Name $VMName -DynamicMemory -MemoryMinimumBytes $MinimumMemoryMB -MB -MemoryMaximumBytes $MaximumMemoryMB -MB -ErrorAction Stop
            Write-Verbose "Dynamic memory configured successfully."
        }
        
        # Create and attach the VHD
        Write-Verbose "Creating virtual hard disk at $VHDPath with size $($VHDSizeBytes/1GB) GB..."
        New-VHD -Path $VHDPath -SizeBytes $VHDSizeBytes -Dynamic -ErrorAction Stop
        Write-Verbose "Virtual hard disk created successfully."
        
        Write-Verbose "Attaching VHD to the VM..."
        Add-VMHardDiskDrive -VMName $VMName -Path $VHDPath -ErrorAction Stop
        Write-Verbose "VHD attached successfully."
        
        # Attach the ISO
        if (-not (Test-Path -Path $ISOPath)) {
            Write-Error "ISO file not found at path: $ISOPath"
            return $false
        }
        
        Write-Verbose "Attaching ISO file from $ISOPath..."
        Add-VMDvdDrive -VMName $VMName -Path $ISOPath -ErrorAction Stop
        Write-Verbose "ISO attached successfully."
        
        # Configure boot order (boot from DVD first)
        Write-Verbose "Configuring boot order to boot from DVD first..."
        $dvdDrive = Get-VMDvdDrive -VMName $VMName
        Set-VMFirmware -VMName $VMName -FirstBootDevice $dvdDrive -ErrorAction Stop
        Write-Verbose "Boot order configured successfully."
        
        # Enable Secure Boot (with Microsoft template)
        Write-Verbose "Enabling Secure Boot with Microsoft template..."
        Set-VMFirmware -VMName $VMName -EnableSecureBoot On -SecureBootTemplate MicrosoftWindows -ErrorAction Stop
        Write-Verbose "Secure Boot enabled successfully."
        
        # Start the VM
        Write-Verbose "Starting the virtual machine..."
        Start-VM -Name $VMName -ErrorAction Stop
        Write-Verbose "Virtual machine started successfully."
        
        Write-Host "Virtual machine '$VMName' created and started successfully!" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to create VM: $_"
        return $false
    }
}

#EndRegion Functions

#Region Main Script Execution

# Display script banner
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   Hyper-V Lab Environment Setup Script      " -ForegroundColor Cyan
Write-Host "   Author: Chris Ibold                       " -ForegroundColor Cyan
Write-Host "   Company: Comet Consulting Group           " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Verify script is running with admin privileges
if (-not (Test-AdminPrivileges)) {
    Write-Error "This script requires administrative privileges. Please run PowerShell as Administrator and try again."
    exit 1
}

# Convert GB values to bytes/MB for internal use
$memoryStartupMB = $MemoryStartupGB * 1024
$vhdSizeBytes = $VHDSizeGB * 1GB
$maximumMemoryMB = $MaximumMemoryGB * 1024

# Configure WinRM
Write-Host "Step 1: Configuring WinRM for remote management..." -ForegroundColor Yellow
$winRMResult = Initialize-WinRM -IPAddress $VMIP
if (-not $winRMResult) {
    Write-Warning "WinRM configuration encountered issues. The script will continue, but remote management might not work properly."
}

# Install Hyper-V if needed
Write-Host "Step 2: Ensuring Hyper-V is installed..." -ForegroundColor Yellow
$hyperVResult = Install-HyperVFeature
if (-not $hyperVResult) {
    Write-Error "Failed to ensure Hyper-V is installed. Please restart your computer if prompted and run the script again."
    exit 1
}

# Create virtual network
Write-Host "Step 3: Setting up virtual network..." -ForegroundColor Yellow
$networkResult = New-VirtualNetwork -SwitchName $SwitchName -SubnetPrefix $Subnet -GatewayIP $GatewayIP
if (-not $networkResult) {
    Write-Error "Failed to set up the virtual network. Please check the error message and try again."
    exit 1
}

# Create the VM
Write-Host "Step 4: Creating the virtual machine..." -ForegroundColor Yellow
$vmResult = New-HyperVLabVM -VMName $VMName `
    -MemoryStartupMB $memoryStartupMB `
    -VHDPath $VHDPath `
    -VHDSizeBytes $vhdSizeBytes `
    -SwitchName $SwitchName `
    -VMPath $VMPath `
    -ISOPath $ISOPath `
    -ProcessorCount $CPUCount `
    -EnableDynamicMemory $EnableDynamicMemory `
    -MinimumMemoryMB $MinimumMemoryMB `
    -MaximumMemoryMB $maximumMemoryMB

if (-not $vmResult) {
    Write-Error "Failed to create the virtual machine. Please check the error message and try again."
    exit 1
}

# Display completion message
Write-Host "=============================================" -ForegroundColor Green
Write-Host "   Hyper-V Lab Environment Setup Complete    " -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "VM Name: $VMName" -ForegroundColor Yellow
Write-Host "VM IP Address: $VMIP" -ForegroundColor Yellow
Write-Host "Hyper-V Switch: $SwitchName" -ForegroundColor Yellow
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Magenta
Write-Host "1. Install the OS through the Hyper-V console" -ForegroundColor White
Write-Host "2. Configure the VM with the IP address: $VMIP" -ForegroundColor White
Write-Host "3. Run the following to connect via PowerShell once the OS is installed:" -ForegroundColor White
Write-Host "`$pass = ConvertTo-SecureString 'YourPassword' -AsPlainText -Force" -ForegroundColor Cyan
Write-Host "`$credential = New-Object System.Management.Automation.PSCredential ('Administrator', `$pass)" -ForegroundColor Cyan
Write-Host "`$session = New-PSSession -ComputerName $VMIP -Credential `$credential" -ForegroundColor Cyan
Write-Host "Enter-PSSession `$session" -ForegroundColor Cyan

#EndRegion Main Script Execution