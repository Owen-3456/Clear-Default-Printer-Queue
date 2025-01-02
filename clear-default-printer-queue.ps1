# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to show a message box
function Show-MessageBox {
    param (
        [string]$Message,
        [string]$Title,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon) | Out-Null
}

# Function to get the default printer name
function Get-DefaultPrinter {
    (Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default = TRUE").Name
}

# Function to check if there are any items in the printer queue
function Get-PrintQueue {
    param (
        [string]$PrinterName
    )
    Get-PrintJob -PrinterName $PrinterName -ErrorAction SilentlyContinue
}

# Function to clear the printer queue
function Clear-PrintQueue {
    param (
        [string]$PrinterName
    )
    try {
        Get-PrintQueue -PrinterName $PrinterName | Remove-PrintJob -ErrorAction Stop
        return $true
    }
    catch {
        Write-Output "Failed to clear the print queue. Error: $_"
        return $false
    }
}

# Main script logic
$DefaultPrinter = Get-DefaultPrinter

if (-not $DefaultPrinter) {
    Show-MessageBox -Message "No default printer is set on this system." -Title "Printer Queue Clearer"
    return
}

$PrintQueue = Get-PrintQueue -PrinterName $DefaultPrinter

if (-not $PrintQueue) {
    Show-MessageBox -Message "There are no items in $DefaultPrinter's queue." -Title "Printer Queue Clearer"
    return
}

$Response = [System.Windows.Forms.MessageBox]::Show("The default printer is '$DefaultPrinter'. Do you want to clear this printer's queue?", "Printer Queue Clearer", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($Response -ne [System.Windows.Forms.DialogResult]::Yes) {
    Write-Output "Operation canceled by user."
    return
}

if (Clear-PrintQueue -PrinterName $DefaultPrinter) {
    $PrintQueue = Get-PrintQueue -PrinterName $DefaultPrinter
    if ($PrintQueue) {
        Show-MessageBox -Message "The task was unsuccessful. The print queue for $DefaultPrinter still contains items." -Title "Task Failed"
    } else {
        Show-MessageBox -Message "The print queue for $DefaultPrinter has been successfully cleared." -Title "Printer Queue Clearer"
    }
}