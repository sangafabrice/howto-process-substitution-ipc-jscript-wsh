#Requires -Version 6.1

# Export the function Convert-MarkdownToHtml.
Get-Item -LiteralPath ([System.IO.Path]::ChangeExtension($PSCommandPath, '.ps1')) |
ForEach-Object {
  New-Item -Path "Function:\$(($FunctionName = 'Convert-MarkdownToHtml'))" -Value (Get-Content $_.FullName -Raw)
  Set-Alias -Name $_.BaseName -Value $FunctionName
}