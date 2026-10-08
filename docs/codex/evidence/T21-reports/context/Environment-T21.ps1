Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
$principal=New-Object Security.Principal.WindowsPrincipal($identity)
[ordered]@{
 Task='T21';ObservedAtUtc=[DateTime]::UtcNow.ToString('o')
 ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition
 Process64Bit=[Environment]::Is64BitProcess;OS64Bit=[Environment]::Is64BitOperatingSystem
 OSVersion=[Environment]::OSVersion.Version.ToString()
 Elevated=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
 ExecutionPolicy=(Get-ExecutionPolicy).ToString()
 PolicyScopes=@(Get-ExecutionPolicy -List | ForEach-Object { [ordered]@{Scope=$_.Scope.ToString();Policy=$_.ExecutionPolicy.ToString()} })
 Culture=[Globalization.CultureInfo]::CurrentCulture.Name
} | ConvertTo-Json -Depth 4
