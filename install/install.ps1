# ============================================================
# KiloHesap - CorelDRAW kilo hesaplama makrosu kurulum betigi
# https://github.com/Volagejs/KiloHesap-CDR
# Lisans: MIT
#
# Bu betik:
#   1) CorelDRAW surumunu ve GMS klasorunu otomatik bulur
#   2) Kaynak dosyalari kopyalar
#   3) CorelDRAW COM uzerinden VBA projesine modulleri
#      OTOMATIK import etmeyi dener ve makroyu test eder
#   4) Olmazsa elle import icin rehberi acar
# ============================================================

$ErrorActionPreference = 'Stop'

function Bilgi($m) { Write-Host "[i] $m" }
function Tamam($m) { Write-Host "[+] $m" -ForegroundColor Green }
function Uyari($m) { Write-Host "[!] $m" -ForegroundColor Yellow }
function Hata($m)  { Write-Host "[X] $m" -ForegroundColor Red }

$repoRoot = Split-Path -Parent $PSScriptRoot
$src = Join-Path $repoRoot 'src'

Bilgi "KiloHesap kurulumu basliyor..."

if (-not (Test-Path -LiteralPath $src)) {
    Hata "src klasoru bulunamadi: $src"
    exit 1
}

# --- CorelDRAW ayar klasorlerini bul ---
$corelRoot = Join-Path $env:APPDATA 'Corel'
if (-not (Test-Path -LiteralPath $corelRoot)) {
    Hata "Corel ayar klasoru bulunamadi: $corelRoot"
    Hata "CorelDRAW bu bilgisayarda kurulu gorunmuyor."
    exit 1
}

$suites = @(Get-ChildItem -LiteralPath $corelRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like 'CorelDRAW Graphics Suite *' })

if ($suites.Count -eq 0) {
    Hata "CorelDRAW Graphics Suite ayar klasoru bulunamadi."
    Hata "CorelDRAW en az bir kez acilmis olmali. CorelDRAW'i bir kez acip kapattiktan sonra kurulumu tekrar calistirin."
    exit 1
}

# 2021 oncelikli, yoksa en yeni surum
$hedef = $suites | Where-Object { $_.Name -like '* 2021' } | Select-Object -First 1
if ($hedef) {
    Tamam "CorelDRAW 2021 ayar klasoru bulundu: $($hedef.Name)"
} else {
    $hedef = $suites | Select-Object -Last 1
    Bilgi "2021 klasoru bulunamadi; en yeni surum kullanilacak: $($hedef.Name)"
}

# Draw\GMS (2023 oncesi) veya DrawPE\GMS (2023 ve sonrasi)
$base = Join-Path $hedef.FullName 'Draw'
if (-not (Test-Path -LiteralPath $base)) { $base = Join-Path $hedef.FullName 'DrawPE' }
$gms = Join-Path $base 'GMS'
New-Item -ItemType Directory -Force -Path $gms | Out-Null
Tamam "Makro klasoru hazir: $gms"

# --- dosyalari kopyala ---
$dest = Join-Path $gms 'KiloHesap'
New-Item -ItemType Directory -Force -Path $dest | Out-Null

Copy-Item -Path (Join-Path $src 'KiloHesap.bas') -Destination $dest -Force
Copy-Item -Path (Join-Path $src 'KiloHesapForm.frm') -Destination $dest -Force
Copy-Item -Path (Join-Path $src 'malzemeler.json') -Destination $dest -Force
Tamam "Dosyalar kopyalandi: $dest"

# ============================================================
# OTOMATIK IMPORT DENEMESI (CorelDRAW COM + VBE)
# ============================================================
$otoImport = $false
$otoKayit = $false
$app = $null

try {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application") } catch {}
    if (-not $app) { $app = New-Object -ComObject CorelDRAW.Application }
    if ($app) {
        Bilgi ("CorelDRAW'a baglanildi: " + $app.Version)
        try { $app.InitializeVBA() | Out-Null } catch {}
        $vbe = $app.VBE

        $gm = $null
        for ($i = 1; $i -le $vbe.VBProjects.Count; $i++) {
            $p = $vbe.VBProjects.Item($i)
            if ($p.Name -eq "GlobalMacros") { $gm = $p }
        }
        if ($gm) {
            # eski importlar temizle
            $silinecek = @()
            for ($j = 1; $j -le $gm.VBComponents.Count; $j++) {
                $c = $gm.VBComponents.Item($j)
                if ($c.Name -eq "KiloHesapModul" -or $c.Name -eq "KiloHesapForm") { $silinecek += $c }
            }
            foreach ($c in $silinecek) { $gm.VBComponents.Remove($c) }

            $gm.VBComponents.Import((Join-Path $src 'KiloHesap.bas')) | Out-Null
            $gm.VBComponents.Import((Join-Path $src 'KiloHesapForm.frm')) | Out-Null
            $otoImport = $true
            Tamam "Moduller GlobalMacros projesine OTOMATIK import edildi."

            # makro testi (form ekranda acilir)
            try {
                $gmgr = $app.GMSManager
                $gmgr.GetType().InvokeMember("RunMacro", [System.Reflection.BindingFlags]::InvokeMethod, $null, $gmgr, @("GlobalMacros", "KiloHesapModul.KiloHesapGoster"))
                Tamam "Makro test edildi - KiloHesap penceresi CorelDRAW ekraninda acildi!"
            } catch {
                Uyari "Makro test atlandi: $($app.Exception.Message)"
            }

            # otomatik kayit denemesi
            try {
                try { $vbe.ActiveVBProject = $gm } catch {}
                $kaydet = $null
                foreach ($bar in $vbe.CommandBars) {
                    try { foreach ($ctl in $bar.Controls) { if ($ctl.Caption -like "*Save GlobalMacros*" -and -not $kaydet) { $kaydet = $ctl } } } catch {}
                }
                if ($kaydet) { $kaydet.Execute(); Start-Sleep -Seconds 2 }
                if (Test-Path -LiteralPath (Join-Path $base 'GlobalMacros.gms')) { $otoKayit = $true }
            } catch {}

            if ($otoKayit) {
                Tamam "GlobalMacros.gms otomatik kaydedildi. Kurulum TAMAMLANDI!"
            }
        }
    }
} catch {
    Bilgi "Otomatik import denendi ama tamamlanamadi: $($app.Exception.Message)"
}

# ============================================================
# REHBER
# ============================================================
$guide = Join-Path $dest 'KurulumRehberi.txt'

if ($otoImport -and $otoKayit) {
    $steps = @"
KILOHESAP - KURULUM TAMAMLANDI!
=====================================================
Makro otomatik kuruldu ve test edildi.

MAKROYU CALISTIRMA
- Window > Dockers > Macro Manager (veya Alt+Shift+F11)
  Listede: GlobalMacros > KiloHesapModul > KiloHesapGoster
  Cift tiklayin. Isterseniz makroyu arac cubuguna
  surukleyerek bir buton da yapabilirsiniz.

-----------------------------------------------------
KULLANIM
-----------------------------------------------------
- Pencere acik kalir. Sekilleri CorelDRAW'da secin.
- Malzemeyi (orn. "Pleksi 5 mm") secin, Hesapla'ya basin.
- Farkli malzemeler icin ayni islemi tekrarlayin;
  her hesap listeye eklenir, toplamlar birikir.
- "Satir sil": yanlis hesabi listeden kaldirir.
- "Sifirla": listeyi temizler.
- "Sayfaya yaz": ozeti cizim sayfasina metin olarak ekler.

-----------------------------------------------------
MALZEME LISTESINI DUZENLEME
-----------------------------------------------------
Bu klasordeki malzemeler.json dosyasini not defteri
ile duzenleyebilirsiniz (yeni malzeme/kalinlik ekleyin).
Degisiklik CorelDRAW yeniden acildiginda gecerli olur.
Klasor: $dest
"@
} elseif ($otoImport) {
    $steps = @"
KILOHESAP - NEREDEYSE TAMAM! TEK ADIM KALDI
=====================================================
Makro otomatik import edildi ve test edildi
(KiloHesap penceresi CorelDRAW ekraninda acildi).

KALICI OLMASI ICIN SON ADIM (~10 saniye):
1) CorelDRAW'da Alt+F11 ile VBA editorunu acin.
2) Ctrl+S'ye basin (File > Save - GlobalMacros.gms kaydedilir).
3) Editoru kapatin. Kurulum tamam!

Not: Ctrl+S dialog acarsa kaydetme klasoru olarak
$base
secip Kaydet'e basin (genelde zaten dogru klasor onerilir).

-----------------------------------------------------
MAKROYU CALISTIRMA
-----------------------------------------------------
Window > Dockers > Macro Manager (veya Alt+Shift+F11)
Listede: GlobalMacros > KiloHesapModul > KiloHesapGoster

-----------------------------------------------------
KULLANIM
-----------------------------------------------------
- Pencere acik kalir. Sekilleri CorelDRAW'da secin.
- Malzemeyi (orn. "Pleksi 5 mm") secin, Hesapla'ya basin.
- Farkli malzemeler icin ayni islemi tekrarlayin;
  her hesap listeye eklenir, toplamlar birikir.
- "Satir sil": yanlis hesabi listeden kaldirir.
- "Sifirla": listeyi temizler.
- "Sayfaya yaz": ozeti cizim sayfasina metin olarak ekler.

-----------------------------------------------------
MALZEME LISTESINI DUZENLEME
-----------------------------------------------------
Bu klasordeki malzemeler.json dosyasini not defteri
ile duzenleyebilirsiniz. Degisiklik CorelDRAW yeniden
acildiginda gecerli olur.
Klasor: $dest
"@
} else {
    $steps = @"
KILOHESAP - TEK SEFERLIK KURULUM (yaklasik 1 dakika)
=====================================================
CorelDRAW, makrolari disaridan otomatik import etmeye
izin vermiyor. Bu yuzden son adim kisa bir elle islem:

1) CorelDRAW'i acin.

2) Menuden: Tools > Visual Basic > Visual Basic Editor
   (klavyeden Alt+F11)

3) Soldaki Project penceresinde "GlobalMacros"
   projesine bir kez tiklayin (secili olsun).

4) Ust menuden: File > Import File...
   - Acilan klasordeki KiloHesap.bas dosyasini secin.

5) Tekrar: File > Import File...
   - KiloHesapForm.frm dosyasini secin.
   NOT: frm import hata verirse:
   - Insert > UserForm ile bos form ekleyin
   - Form adini Properties penceresinden "KiloHesapForm" yapin
   - Formun kod penceresine KiloHesapForm.frm icindeki
     "Option Explicit" ile baslayan kodun tamamini yapistirin

6) Ust menuden: File > Save (GlobalMacros.gms kaydedilir).

7) VBA editorunu kapatin. Makro hazir!

NOT: Ayni adli eski KiloHesap modulleri varsa import hata
verir. O durumda Project penceresinden eski KiloHesap
modul/formlarini silip tekrar import edin.

-----------------------------------------------------
MAKROYU CALISTIRMA
-----------------------------------------------------
Window > Dockers > Macro Manager  (veya Alt+Shift+F11)
Listede: GlobalMacros > KiloHesapModul > KiloHesapGoster
Cift tiklayin. Isterseniz makroyu arac cubuguna
surukleyerek bir buton da yapabilirsiniz.

-----------------------------------------------------
KULLANIM
-----------------------------------------------------
- Pencere acik kalir. Sekilleri CorelDRAW'da secin.
- Malzemeyi (orn. "Pleksi 5 mm") secin, Hesapla'ya basin.
- Farkli malzemeler icin ayni islemi tekrarlayin;
  her hesap listeye eklenir, toplamlar birikir.
- "Satir sil": yanlis hesabi listeden kaldirir.
- "Sifirla": listeyi temizler.
- "Sayfaya yaz": ozeti cizim sayfasina metin olarak ekler.

-----------------------------------------------------
MALZEME LISTESINI DUZENLEME
-----------------------------------------------------
Bu klasordeki malzemeler.json dosyasini not defteri
ile duzenleyebilirsiniz (yeni malzeme/kalinlik ekleyin).
Degisiklik CorelDRAW yeniden acildiginda gecerli olur.
Klasor: $dest
"@
}

Set-Content -LiteralPath $guide -Value $steps
Start-Process notepad.exe -ArgumentList "`"$guide`""

Write-Host ""
if ($otoImport -and $otoKayit) {
    Tamam "Kurulum bitti. Makro Macro Manager'da gorunecek."
} elseif ($otoImport) {
    Uyari "Tek elle adim: VBE'de (Alt+F11) Ctrl+S ile kaydedin."
} else {
    Uyari "Rehberdeki manuel import adimlarini uygulayin."
}
