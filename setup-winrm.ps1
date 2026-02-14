# Run this script as Administrator on each Windows target machine
# to enable and configure WinRM for Ansible

Write-Host "Setting up WinRM for Ansible..." -ForegroundColor Green

# Enable WinRM (even if already enabled, this ensures it's running)
Write-Host "1. Enabling WinRM service..." -ForegroundColor Cyan
Enable-PSRemoting -Force -ErrorAction SilentlyContinue

# Set up firewall rule for WinRM HTTP (5985) and HTTPS (5986)
Write-Host "2. Setting up Windows Firewall rules..." -ForegroundColor Cyan
netsh advfirewall firewall add rule name="WinRM HTTP" dir=in action=allow protocol=tcp localport=5985 enable=yes
netsh advfirewall firewall add rule name="WinRM HTTPS" dir=in action=allow protocol=tcp localport=5986 enable=yes

# Configure WinRM settings for larger payloads and longer timeouts
Write-Host "3. Configuring WinRM for Ansible..." -ForegroundColor Cyan
winrm set winrm/config/winrs '@{MaxMemoryPerShellMB=1024}'
winrm set winrm/config '@{MaxTimeoutms=600000}'
winrm set winrm/config/service '@{MaxConcurrentOperations=100}'

# Enable Basic authentication if needed (make sure users exist locally)
Write-Host "4. Enabling authentication methods..." -ForegroundColor Cyan
winrm set winrm/config/service/auth '@{Basic=true}'
winrm set winrm/config/service '@{AllowUnencrypted=true}'

# Set PowerShell execution policy to allow scripts
Write-Host "5. Setting PowerShell execution policy..." -ForegroundColor Cyan
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force

# Verify WinRM is running
Write-Host "6. Verifying WinRM service is running..." -ForegroundColor Cyan
$service = Get-Service WinRM
if ($service.Status -eq 'Running') {
    Write-Host "✓ WinRM service is running" -ForegroundColor Green
} else {
    Write-Host "✗ WinRM service is NOT running. Starting it now..." -ForegroundColor Yellow
    Start-Service WinRM
}

# Test WinRM connectivity locally
Write-Host "`n7. Testing local WinRM connectivity..." -ForegroundColor Cyan
try {
    Test-WSMan -ComputerName localhost -ErrorAction Stop | Out-Null
    Write-Host "✓ WinRM is responding to local requests" -ForegroundColor Green
} catch {
    Write-Host "✗ WinRM connectivity test failed: $_" -ForegroundColor Red
}

# Verify administrator account exists
Write-Host "`n8. Verifying administrator account..." -ForegroundColor Cyan
try {
    $admin = Get-LocalUser -Name "Administrator" -ErrorAction Stop
    Write-Host "✓ Administrator account found: $($admin.Name)" -ForegroundColor Green
} catch {
    Write-Host "✗ Administrator account not found or error occurred: $_" -ForegroundColor Red
}

Write-Host "`n=== WinRM Setup Complete ===" -ForegroundColor Green
Write-Host "You can now test the connection from your Ansible controller with:" -ForegroundColor Cyan
Write-Host "ansible windows -i inventory.ini -m setup -a 'filter=ansible_date_time'"
