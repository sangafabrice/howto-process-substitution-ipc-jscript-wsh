<#PSScriptInfo .VERSION 0.0.1#>

using namespace System.Windows

Add-Type -AssemblyName PresentationFramework
[MessageBox]::Show($args[0].Substring(1).Remove($args[0].Length - 2).Replace('*', '"').Replace('^', "`n").Trim(), 'Convert to HTML', $args[1], $args[2])