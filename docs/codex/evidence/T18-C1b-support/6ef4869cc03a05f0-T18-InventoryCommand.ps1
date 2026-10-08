$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
try { $admin=(New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } finally { $identity.Dispose() }
$record=[ordered]@{observed_at_utc=[DateTime]::UtcNow.ToString("o");shell_version=$PSVersionTable.PSVersion.ToString();process_64_bit=[Environment]::Is64BitProcess;administrator=$admin;policy=(Get-ExecutionPolicy).ToString();scopes=@(Get-ExecutionPolicy -List | ForEach-Object {"$($_.Scope)=$($_.ExecutionPolicy)"});os_version=[Environment]::OSVersion.Version.ToString();filesystem=(New-Object IO.DriveInfo("C:\")).DriveFormat;process_execution_policy_argument=$false;application_or_native_test_claim=$false}
$record | ConvertTo-Json -Depth 5
