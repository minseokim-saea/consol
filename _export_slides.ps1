$ErrorActionPreference = "Stop"
$pptPath = "C:\패키지프로그램\연결재무보고시스템_임원보고용.pptx"
$outDir  = "C:\패키지프로그램\_slide_preview"

if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
# Visible 속성 설정 시도 - 일부 PPT 버전에서 강제 필요
try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
try {
    $pres = $ppt.Presentations.Open($pptPath, $true, $false, $false)
} catch {
    $pres = $ppt.Presentations.Open($pptPath)
}

# 각 슬라이드를 PNG로 export (가로 1600px)
$pres.SaveAs($outDir, 18)  # 18 = ppSaveAsPNG
$pres.Close()
$ppt.Quit()

Write-Output "Done: $outDir"
Get-ChildItem $outDir | Select-Object Name, Length | Format-Table
