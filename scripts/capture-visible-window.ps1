param(
  [Parameter(Mandatory = $true)]
  [string]$ProcessName,

  [string]$TitleContains = '',

  [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class VisibleWindowCapture {
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
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd);
}
"@

if (-not $OutputPath) {
  $OutputPath = Join-Path (Get-Location) "logs\\${ProcessName}-visible.png"
}

$candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
  $_.ProcessName -ieq $ProcessName -and $_.MainWindowHandle -ne 0
}

if ($TitleContains) {
  $candidates = $candidates | Where-Object { $_.MainWindowTitle -like "*$TitleContains*" }
}

$windows = @()
foreach ($candidate in $candidates) {
  if (-not [VisibleWindowCapture]::IsWindowVisible($candidate.MainWindowHandle)) { continue }
  if ([VisibleWindowCapture]::IsIconic($candidate.MainWindowHandle)) { continue }

  $rect = New-Object VisibleWindowCapture+RECT
  if (-not [VisibleWindowCapture]::GetWindowRect($candidate.MainWindowHandle, [ref]$rect)) { continue }

  $width = $rect.Right - $rect.Left
  $height = $rect.Bottom - $rect.Top
  if ($width -le 0 -or $height -le 0) { continue }

  $windows += [PSCustomObject]@{
    Process = $candidate
    Rect = $rect
    Width = $width
    Height = $height
    Area = $width * $height
  }
}

$selected = $windows |
  Sort-Object -Property @(
    @{ Expression = { $_.Area }; Descending = $true },
    @{ Expression = { $_.Process.StartTime }; Descending = $true }
  ) |
  Select-Object -First 1

if (-not $selected) {
  throw "No visible non-minimized window found for process '$ProcessName'."
}

$bmp = New-Object System.Drawing.Bitmap $selected.Width, $selected.Height
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen($selected.Rect.Left, $selected.Rect.Top, 0, 0, $bmp.Size)

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) {
  New-Item -ItemType Directory -Path $dir | Out-Null
}

$bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bmp.Dispose()

Write-Output $OutputPath
