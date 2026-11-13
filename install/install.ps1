# ============================================================
# KiloHesap - CorelDRAW kilo hesaplama makrosu kurulum betigi
# https://github.com/Volagejs/KiloHesap-CDR
# Lisans: MIT
# ============================================================

$ErrorActionPreference = 'Stop'

function Bilgi($m) { Write-Host "[i] $m" }
function Tamam($m) { Write-Host "[+] $m" -ForegroundColor Green }
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

if (Test-Path -LiteralPath (Join-Path $gms 'KiloHesap.gms')) {
    Bilgi "KiloHesap.gms zaten kurulu. VBA editorunde modulleri tekrar import edip kaydederek guncelleyebilirsiniz."
}

# --- kurulum rehberi ---
$guide = Join-Path $dest 'KurulumRehberi.txt'
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

Set-Content -LiteralPath $guide -Value $steps

Write-Host ""
Bilgi "Son adim VBA editorunde import gerektiriyor (CorelDRAW siniri)."
Bilgi "Kurulum rehberi notepad ile aciliyor..."
Start-Process notepad.exe -ArgumentList "`"$guide`""
Start-Process explorer.exe -ArgumentList "/select,`"$dest`""

Write-Host ""
Tamam "Kurulum dosyalari hazir. Rehberdeki adimlari uygulayin."
