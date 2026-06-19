if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$dest = "C:\Program Files (x86)\Adobe\Adobe Illustrator CS4\Presets\en_US\Scripts"
Get-ChildItem -Path $PSScriptRoot -Exclude ".git","deploy.ps1","deploy.sh","readme.md" | Copy-Item -Destination $dest -Force
Write-Host "Done."
