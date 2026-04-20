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

public static class BackgroundWindowCapture {
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
  public static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd);
}
"@

if (-not $OutputPath) {
  $OutputPath = Join-Path (Get-Location) "logs\\${ProcessName}-background.png"
}

$candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
  $_.ProcessName -ieq $ProcessName -and $_.MainWindowHandle -ne 0
}

if ($TitleContains) {
  $candidates = $candidates | Where-Object { $_.MainWindowTitle -like "*$TitleContains*" }
}

$windows = @()
foreach ($candidate in $candidates) {
  if (-not [BackgroundWindowCapture]::IsWindowVisible($candidate.MainWindowHandle)) {
    continue
  }

  $candidateRect = New-Object BackgroundWindowCapture+RECT
  if (-not [BackgroundWindowCapture]::GetWindowRect($candidate.MainWindowHandle, [ref]$candidateRect)) {
    continue
  }

  $candidateWidth = $candidateRect.Right - $candidateRect.Left
  $candidateHeight = $candidateRect.Bottom - $candidateRect.Top
  if ($candidateWidth -le 0 -or $candidateHeight -le 0) {
    continue
  }

  $windows += [PSCustomObject]@{
    Process = $candidate
    Rect = $candidateRect
    Width = $candidateWidth
    Height = $candidateHeight
    Area = $candidateWidth * $candidateHeight
    IsMinimized = [BackgroundWindowCapture]::IsIconic($candidate.MainWindowHandle)
    SortMinimized = if ([BackgroundWindowCapture]::IsIconic($candidate.MainWindowHandle)) { 1 } else { 0 }
  }
}

$selected = $windows |
  Sort-Object -Property @(
    @{ Expression = { $_.SortMinimized }; Descending = $false },
    @{ Expression = { $_.Area }; Descending = $true },
    @{ Expression = { $_.Process.StartTime }; Descending = $true }
  ) |
  Select-Object -First 1

if (-not $selected) {
  throw "No visible background-capturable window found for process '$ProcessName'."
}

$bmp = New-Object System.Drawing.Bitmap $selected.Width, $selected.Height
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$hDC = $graphics.GetHdc()

try {
  $ok = [BackgroundWindowCapture]::PrintWindow($selected.Process.MainWindowHandle, $hDC, 0)
} finally {
  $graphics.ReleaseHdc($hDC)
}

if (-not $ok) {
  $graphics.Dispose()
  $bmp.Dispose()
  throw "PrintWindow failed for process '$ProcessName'."
}

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) {
  New-Item -ItemType Directory -Path $dir | Out-Null
}

$bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bmp.Dispose()

Write-Output $OutputPath
