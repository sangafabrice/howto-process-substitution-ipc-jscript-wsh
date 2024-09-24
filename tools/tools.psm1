<#PSScriptInfo .VERSION 1.0.0#>

Get-Item "$PSScriptRoot\*.ps1" | ForEach-Object { New-Item -Path "Function:\" -Name $_.BaseName -Value (Get-Content $_.FullName -Raw) }
Export-ModuleMember -Function *