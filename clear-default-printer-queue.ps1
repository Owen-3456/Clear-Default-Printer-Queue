# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Get the default printer name
$DefaultPrinter = (Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default = TRUE").Name

if (-not $DefaultPrinter) {
    [System.Windows.Forms.MessageBox]::Show("No default printer is set on this system.", "Printer Queue Clearer", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    return
}

# Check if there are any items in the printer queue
$PrintQueue = Get-PrintJob -PrinterName $DefaultPrinter -ErrorAction SilentlyContinue

if (-not $PrintQueue) {
    [System.Windows.Forms.MessageBox]::Show("There are no items in $DefaultPrinter's queue.", "Printer Queue Clearer", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    return
}

# Ask user if this is the correct printer using a pop-up box
$Response = [System.Windows.Forms.MessageBox]::Show("The default printer is '$DefaultPrinter'. Do you want to clear this printer's queue?", "Printer Queue Clearer", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($Response -ne [System.Windows.Forms.DialogResult]::Yes) {
    Write-Output "Operation canceled by user."
    return
}

if ($Response -ne "Y") {
    Write-Output "Operation canceled by user."
    return
}

# Clear the printer queue
try {
    $PrintQueue | Remove-PrintJob -ErrorAction Stop
    
    $PrintQueue = Get-PrintJob -PrinterName $DefaultPrinter -ErrorAction SilentlyContinue

    if (-not $PrintQueue) {
        [System.Windows.Forms.MessageBox]::Show("The print queue for $DefaultPrinter has been successfully cleared.", "Printer Queue Clearer", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
}
catch {
    Write-Output "Failed to clear the print queue. Error: $_"
}
