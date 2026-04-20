$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$signature = @'
using System;
using System.Runtime.InteropServices;

public static class NativeWindowCapture {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }

  [DllImport("user32.dll")]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

  [DllImport("user32.dll")]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
}
'@

Add-Type -TypeDefinition $signature

$outputPath = if ($args.Length -gt 0) { $args[0] } else { Join-Path (Get-Location) 'logs\\electron-window.png' }

$proc = Get-Process | Where-Object {
  $_.ProcessName -eq 'electron' -and $_.MainWindowHandle -ne 0
} | Sort-Object StartTime -Descending | Select-Object -First 1

if (-not $proc) {
  throw 'No Electron window with a main handle is currently running.'
}

[NativeWindowCapture]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
Start-Sleep -Milliseconds 400

$rect = New-Object NativeWindowCapture+RECT
if (-not [NativeWindowCapture]::GetWindowRect($proc.MainWindowHandle, [ref]$rect)) {
  throw 'Failed to get Electron window bounds.'
}

$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top

if ($width -le 0 -or $height -le 0) {
  throw "Invalid Electron window size: ${width}x${height}"
}

$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)

$dir = Split-Path -Parent $outputPath
if ($dir -and -not (Test-Path $dir)) {
  New-Item -ItemType Directory -Path $dir | Out-Null
}

$bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()

Write-Output $outputPath
