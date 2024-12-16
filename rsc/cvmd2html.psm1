#Requires -Version 6.1
using namespace System.IO

Function Convert-MarkdownToHtml {
  <#
  .SYNOPSIS
  Convert a Markdown file to an HTML file.
  .DESCRIPTION
  The script converts the specified Markdown file to an HTML file. It shows an overwrite prompt when the HTML already exists.
  .PARAMETER MarkdownPath
  Specifies the path of an existing .md Markdown file.
  .PARAMETER HtmlFilePath
  Specifies the path of the output HTML file.
  By default, the output HTML file has the same parent directory and base name as the input Markdown file.
  .PARAMETER OverWrite
  Specifies that the output file should be overriden.
  .EXAMPLE
  "Here's the link to the [team session](https://fromthetechlab.blogspot.com)." > .\Readme.md
  PS> Convert-MarkdownToHtml .\Readme.md
  PS> Get-Content .\Readme.html
  <p>Here's the link to the <a href="https://fromthetechlab.blogspot.com">team session</a>.</p>
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory, Position=0)]
    [ValidateScript({Test-Path $_ -PathType Leaf}, ErrorMessage = 'The input file "{0}" is not found.')]
    [ValidatePattern('\.md$', ErrorMessage = 'The extension of "{0}" is invalid. ".md" is required.')]
    [string] $MarkdownPath,
    [Parameter(Position=1)]
    [ValidatePattern('\.html?$', ErrorMessage = 'The extension of "{0}" is invalid. ".htm" or "html" is required.')]
    [string] $HtmlPath = [Path]::ChangeExtension($MarkdownPath, '.html'),
    [switch] $OverWrite
  )

  # Handle exceptions when the output path is
  # an already existing HTML file or a directory.
  If (Test-Path $HtmlPath -PathType Leaf) {
    If (-not $OverWrite) {
      # Private variable to copy LASTEXITCODE value.
      $Private:LASTEXITCODE = $LASTEXITCODE
      Write-Host ('The file "{0}" already exists.' -f $HtmlPath)
      Write-Host 'Do you want to overwrite it?'
      C:\Windows\System32\choice.exe /C YN /N /M '[Y]es [N]o: '
      # Exit when the user chooses No.
      If ($Global:LASTEXITCODE -eq 2) {
        # Remove impact of choice.exe on LASTEXITCODE.
        $Global:LASTEXITCODE = $Private:LASTEXITCODE
        Return
      }
      # Remove impact of choice.exe on LASTEXITCODE.
      $Global:LASTEXITCODE = $Private:LASTEXITCODE
    }
  } ElseIf (Test-Path $HtmlPath) {
    Throw ('"{0}" cannot be overwritten because it is a directory.' -f $HtmlPath)
  }
  # Conversion from Markdown to HTML.
  (ConvertFrom-Markdown $MarkdownPath -ErrorAction Stop).Html | Out-File $HtmlPath
}

Set-Alias -Name cvmd2html -Value Convert-MarkdownToHtml