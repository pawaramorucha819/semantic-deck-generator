# 生成した .pptx を PNG に書き出して目視確認するための補助スクリプト（Windows + PowerPoint 必須）
#
#   powershell -File scripts/preview-pptx.ps1 `
#     -Pptx examples/generated/inventory-system.pptx `
#     -OutDir .preview
#
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [string]$OutDir = ".preview",
  [int]$Width = 1600,
  [int]$Height = 900
)

$ErrorActionPreference = "Stop"

$Pptx = (Resolve-Path $Pptx).Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path
Get-ChildItem -Path $OutDir -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force

$app = New-Object -ComObject PowerPoint.Application
try {
  # ReadOnly / Untitled=false / WithWindow=false で開く
  $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
  try {
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $target = Join-Path $OutDir ("slide-{0:d2}.png" -f $i)
      $pres.Slides.Item($i).Export($target, "PNG", $Width, $Height)
      Write-Output "exported $target"
    }
  } finally {
    $pres.Close()
  }
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
