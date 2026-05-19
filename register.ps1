$dllPath = "$PSScriptRoot\bin\Debug\sign-templater.dll"
$vstoPath = "$PSScriptRoot\bin\Debug\sign-templater.vsto"

if (-not (Test-Path $vstoPath)) {
    Write-Host "Error: .vsto file not found. Build project in Visual Studio first!" -ForegroundColor Red
    exit
}

$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\sign-templater"

if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

Set-ItemProperty -Path $regPath -Name "Description" -Value "Sign Templater VSTO Add-in"
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Sign Templater"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "$vstoPath|vstolocal" -PropertyType String

Write-Host "Success: Registered 'sign-templater' for Word." -ForegroundColor Green
Write-Host "Now restart Word."
