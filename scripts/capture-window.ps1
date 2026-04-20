param(
  [Parameter(Mandatory = $true)]
  [string]$ProcessName,

  [string]$TitleContains = '',

  [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$signature = @'
using System;
using System.Runtime.InteropServices;

public static class NativeWindowCaptureAny {
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

if (-not $OutputPath) {
  $OutputPath = Join-Path (Get-Location) "logs\\${ProcessName}-window.png"
}

$candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
  $_.ProcessName -ieq $ProcessName -and $_.MainWindowHandle -ne 0
}

if ($TitleContains) {
  $candidates = $candidates | Where-Object { $_.MainWindowTitle -like "*$TitleContains*" }
}

$windows = @()
foreach ($candidate in $candidates) {
  $candidateRect = New-Object NativeWindowCaptureAny+RECT
  if (-not [NativeWindowCaptureAny]::GetWindowRect($candidate.MainWindowHandle, [ref]$candidateRect)) {
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
  }
}

$selected = $windows | Sort-Object -Property @(
  @{ Expression = { $_.Area }; Descending = $true },
  @{ Expression = { $_.Process.StartTime }; Descending = $true }
) | Select-Object -First 1
$proc = $selected.Process

if (-not $proc) {
  throw "No window found for process '$ProcessName'."
}

[NativeWindowCaptureAny]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
Start-Sleep -Milliseconds 500

$rect = $selected.Rect
$width = $selected.Width
$height = $selected.Height

if ($width -le 0 -or $height -le 0) {
  throw "Invalid window size: ${width}x${height}"
}

$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) {
  New-Item -ItemType Directory -Path $dir | Out-Null
}

$bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()

Write-Output $OutputPath
