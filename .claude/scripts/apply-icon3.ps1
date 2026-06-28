# Get the actual shortcut file name from the directory
$lnkDir = "C:\Users\12969\Desktop"
$lnkName = (Get-ChildItem $lnkDir -Filter "*.lnk" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "Claude*" } | Select-Object -First 1).Name
Write-Host "Found lnk: [$lnkName]"
Write-Host "Bytes: $([System.Text.Encoding]::UTF8.GetBytes($lnkName) -join ',')"

# Build path using the actual name
$lnkPath = Join-Path $lnkDir $lnkName
$icoPath = "D:\TREA文件\.claude\scripts\claude-code.ico"

Write-Host "Full lnk path: [$lnkPath]"
Write-Host "Test-Path: $(Test-Path $lnkPath)"

# Update shortcut icon
$ws = New-Object -ComObject WScript.Shell
$shortcut = $ws.CreateShortcut($lnkPath)
if ($shortcut -eq $null) {
  Write-Host "ERROR: shortcut is null"
  exit 1
}
Write-Host "Current TargetPath: $($shortcut.TargetPath)"
Write-Host "Current IconLocation: $($shortcut.IconLocation)"

$shortcut.IconLocation = "$icoPath,0"
$shortcut.Save()
Write-Host "New IconLocation set to: $icoPath,0"

# Re-verify
$shortcut2 = $ws.CreateShortcut($lnkPath)
Write-Host "Verify IconLocation: $($shortcut2.IconLocation)"

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut2) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
Write-Host "Done!"
