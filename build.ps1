<#PSScriptInfo .VERSION 1.0.1#>

using namespace System.IO
using namespace System.Runtime.InteropServices
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.js','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })
  Set-ProjectVersion $PSScriptRoot
  Remove-Module tools
}