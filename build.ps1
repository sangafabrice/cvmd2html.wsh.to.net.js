#Requires -Version 6.1
using namespace System.IO
using namespace System.Management.Automation
[CmdletBinding()]
Param ()

Function Build-MarkdownToHtmlShortcut {
  <#
  .SYNOPSIS
  Build the project around one or more executables.
  .DESCRIPTION
  This function creates the binary and resource files that are ready for distribution.
  .OUTPUTS
  The executable path string.
  #>
  [CmdletBinding()]
  Param ()

  $HostColorArgs = @{
    ForegroundColor = 'Black'
    BackgroundColor = 'Green'
    NoNewline = $True
  }
  # Set the extension of the file with base name Convert-MarkdownToHtml and return full path.
  Function Private:Set-ConvertMd2HtmlExtension([string] $Extension) {
    Return "$PSScriptRoot\Convert-MarkdownToHtml$Extension"
  }
  Try {
    Remove-Item ($ConvertExe = Set-ConvertMd2HtmlExtension '.exe') -ErrorAction Stop
  }
  Catch [ItemNotFoundException] {
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
  }
  Catch {
    $HostColorArgs.BackgroundColor = 'Red'
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
    Return
  }
  # Import the dependency libraries.
  $ScriptletDllPath = "$PSScriptRoot\Interop.Scriptlet.dll";
  $SWbemDllPath = "$PSScriptRoot\Interop.WbemScripting.dll";
  $WshDllPath = "$PSScriptRoot\Interop.IWshRuntimeLibrary.dll";
  & "$PSScriptRoot\TlbImp.exe" /nologo /silent 'C:\Windows\System32\scrobj.dll'  /out:$ScriptletDllPath /namespace:Scriptlet
  & "$PSScriptRoot\TlbImp.exe" /nologo /silent 'C:\Windows\System32\wshom.ocx'  /out:$WshDllPath /namespace:IWshRuntimeLibrary
  & "$PSScriptRoot\TlbImp.exe" /nologo /silent 'C:\Windows\System32\wbem\wbemdisp.tlb' /out:$SWbemDllPath /namespace:WbemScripting
  # Compile the launcher script into an .exe file of the same base name.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:windir\Microsoft.NET\Framework$(If ([Environment]::Is64BitOperatingSystem) { '64' })\v4.0.30319\;$Env:Path"
  jsc.exe /nologo /target:$(($IsContinue = $DebugPreference -eq 'Continue') ? 'exe':'winexe') /reference:$WshDllPath /reference:$SWbemDllPath /reference:$ScriptletDllPath /out:$(($ConvertExe = Set-ConvertMd2HtmlExtension '.exe')) $(Set-ConvertMd2HtmlExtension '.js')
  $Env:Path = $EnvPath
  If ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs
    (Get-Item $ConvertExe).VersionInfo | Format-List * -Force
  }
  $StartArgs = @{
    FilePath = $ConvertExe
    NoNewWindow = $True
    Wait = $True
  }
  Start-Process @StartArgs -ArgumentList '/Unset'
  If ($IsContinue) {
    Start-Sleep 5
  }
  Start-Process @StartArgs -ArgumentList '/Set'
}

Build-MarkdownToHtmlShortcut -Debug:$($DebugPreference -eq 'Continue')