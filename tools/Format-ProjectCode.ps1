<#PSScriptInfo .VERSION 1.0.0#>

[CmdletBinding()]
param ([string[]] $LiteralPath)
Get-ChildItem $LiteralPath -Recurse | ForEach-Object {
  $content = @(Get-Content $_.FullName).TrimEnd() -join [Environment]::NewLine
  Set-Content $_.FullName $content -NoNewLine
}