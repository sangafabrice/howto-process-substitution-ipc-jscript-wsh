<#PSScriptInfo .VERSION 1.0.0#>

using namespace System.IO
using namespace System.Runtime.InteropServices
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.js','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })
  Remove-Module tools
}