
<#
    Description: Import Addresses from CSV with Logging and Module Check
#>

# ---------------------------------------------------------------------------
# 1. การตั้งค่าตัวแปร (CONFIG)
# ---------------------------------------------------------------------------
$CsvFilePath = ".\uploads\addresses.csv"  # <-- แก้ไข path ไฟล์ CSV ของคุณที่นี่
$LogFolder = ".\logs"          # <-- โฟลเดอร์ที่จะเก็บ Log
$LogFile = "$LogFolder\AddressLog_$(Get-Date -Format 'yyyyMMdd-HHmm').txt"

# สร้าง Folder Log หากยังไม่มี
if (!(Test-Path -Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder | Out-Null
}

# ฟังก์ชันสำหรับเขียน Log
function Write-Log {
    Param ([string]$Message, [string]$Type = "INFO")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogContent = "[$TimeStamp] [$Type] $Message"
    Add-Content -Path $LogFile -Value $LogContent
    Write-Host $LogContent -ForegroundColor $(IF ($Type -eq "ERROR") { "Red" } elseif ($Type -eq "WARNING") { "Yellow" } else { "Green" })
}

# ---------------------------------------------------------------------------
# 2. ตรวจสอบและติดตั้ง Module (MODULE CHECK)
# ---------------------------------------------------------------------------
Write-Log "Starting Script..."
try {
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Module 'ExchangeOnlineManagement' not found. Installing..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Log "Module installed successfully."
    }
    else {
        Write-Log "Module 'ExchangeOnlineManagement' is already installed."
    }
}
catch {
    Write-Log "CRITICAL ERROR: Failed to install module. Details: $_" "ERROR"
    Exit
}


# ---------------------------------------------------------------------------
# 3. เชื่อมต่อ Exchange Online (CONNECT)
# ---------------------------------------------------------------------------
try {
    # เช็คว่าต่ออยู่แล้วหรือยัง ถ้ายังให้ connect
    if (!([bool](Get-ConnectionInformation -ErrorAction SilentlyContinue))) {
        Write-Log "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        Write-Log "Connected successfully."
    }
}
catch {
    Write-Log "CRITICAL ERROR: Failed to connect to Exchange Online. Details: $_" "ERROR"
    Exit
}

# ---------------------------------------------------------------------------
# 4. เริ่มกระบวนการ Import/Update (PROCESS)
# ---------------------------------------------------------------------------

# ตรวจสอบว่าไฟล์ CSV มีอยู่จริงไหม
if (!(Test-Path $CsvFilePath)) {
    # ในกรณีต้องการเลือกไฟล์ผ่าน Dialog
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = "C:\"
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.Multiselect = $false
    $dialog.Title = "Select the CSV file for Addresses"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $CsvFilePath = $dialog.FileName
    }
    else {
        Write-Log "No CSV file selected. Exiting." "ERROR"
        Exit
    }
}

# อ่านไฟล์ CSV
try {
    $Addresses = Import-Csv $CsvFilePath
    Write-Log "Loaded $( $Addresses.Count ) Addresses from CSV."
}
catch {
    Write-Log "Error reading CSV file. Details: $_" "ERROR"
    Exit
}


# ต้องการอัปเดต Address ที่มีอยู่แล้วหรือไม่
$Is_skip_confirmation = Read-Host "Do you want to updating existing addresses? (Y/N) Default(Y) "
$Skip_update = $Is_skip_confirmation -ne "Y" -and $Is_skip_confirmation -ne "y"
if ($Skip_update) {
    Write-Log "User chose to skip updating existing addresses."
}

# วนลูปเพิ่ม Address
foreach ($Row in $Addresses) {
    $Name = $Row."Display Name"
    $Rule = $Row.Rule
    $Value = $Row.Value
    try {
        $ExistingAddress = Get-AddressList -Identity $Name -ErrorAction SilentlyContinue
        if ($ExistingAddress) {
            
            if ($Skip_update) {
                Write-Log "Address found: $Name. Skipping update as per user choice."
                continue
            }
         
            # --- กรณีมีอยู่แล้ว ให้ UPDATE ---
            Write-Log "Address found: $Name. Updating..."
            IF (![string]::IsNullOrEmpty($Rule) -and ![string]::IsNullOrEmpty($Value) ) {
                # Build the recipient filter based on the rule and value
                $RecipientFilter = "$Rule -like '$Value'"
                # Update the existing address list
                Set-AddressList -Identity $Name -RecipientFilter $RecipientFilter -ErrorAction Stop
                Write-Log "Updated Rule for Address: $Name with filter: $RecipientFilter"
            }
            ELSE {
                Write-Log "No Rule or Value provided for update for Address: $Name" "WARNING"
            }

        }
        else {
            # --- กรณีไม่มี ให้สร้างใหม่ ---
            Write-Log "Address not found: $Name. Creating new Address..."
            IF (![string]::IsNullOrEmpty($Rule) -and ![string]::IsNullOrEmpty($Value)) {
                # Build the recipient filter based on the rule and value
                $RecipientFilter = "$Rule -like '$Value'"
                
                # Create new address list
                New-AddressList -Name $Name -RecipientFilter $RecipientFilter -ErrorAction Stop
                Write-Log "Created new Address: $Name with filter: $RecipientFilter"
            }
            ELSE {
                Write-Log "No Rule or Value provided for creation for Address: $Name" "WARNING"
            }
        }
    }
    catch {
        Write-Log "Error processing address '$Name': $($_.Exception.Message)" "ERROR"
    }
}

# ---------------------------------------------------------------------------
# 5. สรุปผลการทำงานและปิดการเชื่อมต่อ (SUMMARY & CLEANUP)
# ---------------------------------------------------------------------------
Write-Log "Address import/update process completed."
Write-Log "Check the log file at: $LogFile"

# ---------------------------------------------------------------------------
# 6. ตัดการเชื่อมต่อ Exchange Online (DISCONNECT)
# ---------------------------------------------------------------------------
$Is_Disconnect = Read-Host "Do you want to disconnect from Exchange Online? (Y/N) Default(N) "
$Disconnect = $Is_Disconnect -eq "Y" -or $Is_Disconnect -eq "y"

try {
    if ($Disconnect) {
        Write-Log "Disconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Disconnected successfully."
    }
    else {
        Write-Log "User chose to remain connected to Exchange Online."
    }
}
catch {
    Write-Log "Error disconnecting from Exchange Online. Details: $_" "ERROR"
}