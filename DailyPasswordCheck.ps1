# =========================================
# DAILY ACTIVE DIRECTORY PASSWORD CHECK
# WITH EXCEL REPORT AND LAST RUN TIMESTAMP
# =========================================

Import-Module ActiveDirectory
Import-Module ImportExcel

# ---------- CONFIGURATION ---------- #

$PASSWORD_CHANGE_DAYS = 45
$EXCEL_PATH = "C:\PasswordChangeReport.xlsx"
$OU_PATH = "OU=_EMPLOYEES," + ([ADSI]"").distinguishedName

# ---------- DATE / TIME STAMP ---------- #

$today = Get-Date
$runTimestamp = Get-Date -Format "MM-dd-yyyy HH:mm:ss"

Write-Host ""
Write-Host "Starting Daily Password Expiration Scan..." -ForegroundColor Yellow
Write-Host "Task Last Executed: $runTimestamp" -ForegroundColor Magenta
Write-Host ""

# ---------- GET USERS ---------- #

$users = Get-ADUser `
    -SearchBase $OU_PATH `
    -Filter * `
    -Properties Description, PasswordNeverExpires

# ---------- BUILD REPORT ---------- #

$report = foreach ($user in $users) {

    if ($user.Description -match "StartDate=(\d{4}-\d{2}-\d{2})") {

        $startDate = Get-Date $matches[1]
        $passwordDueDate = $startDate.AddDays($PASSWORD_CHANGE_DAYS)
        $daysRemaining = ($passwordDueDate - $today).Days

        if ($today -ge $passwordDueDate) {

            $status = "PASSWORD CHANGE REQUIRED"

            Set-ADUser `
                $user.SamAccountName `
                -ChangePasswordAtLogon $true
        }
        else {
            $status = "Not Due Yet"
        }

        [PSCustomObject]@{
            Username = $user.SamAccountName
            StartDate = $startDate.ToString("MM-dd-yyyy")
            PasswordChangeDue = $passwordDueDate.ToString("MM-dd-yyyy")
            DaysRemaining = $daysRemaining
            Status = $status
            LastRunTimestamp = $runTimestamp
        }
    }
}

# ---------- DISPLAY TABLE ---------- #

Write-Host ""
Write-Host "Password Expiration Report" -ForegroundColor Cyan

$report |
    Sort-Object PasswordChangeDue |
    Format-Table Username, StartDate, PasswordChangeDue, DaysRemaining, Status, LastRunTimestamp -AutoSize

# ---------- EXPORT TO EXCEL ---------- #

$report |
    Sort-Object PasswordChangeDue |
    Export-Excel `
        -Path $EXCEL_PATH `
        -WorksheetName "Password Report" `
        -AutoSize `
        -BoldTopRow `
        -FreezeTopRow `
        -TableName "PasswordChangeReport" `
        -ClearSheet

Write-Host ""
Write-Host "Excel report updated: $EXCEL_PATH" -ForegroundColor Green
Write-Host "Task Last Executed: $runTimestamp" -ForegroundColor Magenta
Write-Host ""