@echo off
title KiloHesap Kaldir
echo ============================================
echo  KiloHesap kaldiriliyor...
echo ============================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$suites=Get-ChildItem (Join-Path $env:APPDATA 'Corel') -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -like 'CorelDRAW Graphics Suite *'}; if(-not $suites){Write-Host 'CorelDRAW klasoru bulunamadi.'}; foreach($s in $suites){ foreach($sub in @('Draw','DrawPE')){ $d=Join-Path (Join-Path $s.FullName $sub) 'GMS\KiloHesap'; if(Test-Path $d){ Remove-Item $d -Recurse -Force; Write-Host ('Silindi: '+$d) }; $g=Join-Path (Join-Path $s.FullName $sub) 'GMS\KiloHesap.gms'; if(Test-Path $g){ Remove-Item $g -Force; Write-Host ('Silindi: '+$g) } } }; Write-Host ''; Write-Host 'KiloHesap kaldirildi. VBA editorunde GlobalMacros icindeki KiloHesap modul/formlarini da elle silebilirsiniz.'"
echo.
pause
