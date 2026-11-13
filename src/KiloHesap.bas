Attribute VB_Name = "KiloHesapModul"
Option Explicit

' =============================================================
' KiloHesap - CorelDRAW icin kilo hesaplama makrosu
' Uyumluluk: CorelDRAW 2019-2024 (CorelDRAW 2021 icin hazirlandi)
' https://github.com/Volagejs/KiloHesap-CDR
' Lisans: MIT
'
' Formul:
'   kilo (g) = alan (cm2) x kalinlik (cm) x yogunluk (g/cm3)
' =============================================================

Public Type Malzeme
    ad As String
    kalinlikMM As Double
    yogunluk As Double
End Type

Public gMalzemeler() As Malzeme
Public gMalzemeSayisi As Long

' -------------------------------------------------------------
' Makro giris noktasi. Pencereyi mod (modeless) olarak acar;
' pencere acik kalirken CorelDRAW'da sekil secilebilir.
' -------------------------------------------------------------
Public Sub KiloHesapGoster()
    MalzemeleriYukle
    KiloHesapForm.Show vbModeless
End Sub

' -------------------------------------------------------------
' Malzemeler
' -------------------------------------------------------------
Public Sub MalzemeleriYukle()
    VarsayilanMalzemeler
    Dim yol As String
    yol = JsonYoluBul()
    If Len(yol) > 0 Then JsonOku yol
End Sub

Private Sub VarsayilanMalzemeler()
    gMalzemeSayisi = 11
    ReDim gMalzemeler(1 To gMalzemeSayisi)
    MalzemeAta 1, "Pleksi", 3#, 1.19
    MalzemeAta 2, "Pleksi", 5#, 1.19
    MalzemeAta 3, "Pleksi", 8#, 1.19
    MalzemeAta 4, "Pleksi", 10#, 1.19
    MalzemeAta 5, "Dekota (MDF)", 9#, 0.75
    MalzemeAta 6, "Dekota (MDF)", 14#, 0.75
    MalzemeAta 7, "Dekota (MDF)", 18#, 0.75
    MalzemeAta 8, "Krom", 1#, 7.9
    MalzemeAta 9, "Aluminyum", 2#, 2.7
    MalzemeAta 10, "Krom Yan Sac", 0.6, 7.9
    MalzemeAta 11, "Aluminyum Yan Sac", 0.7, 2.7
End Sub

Private Sub MalzemeAta(ByVal idx As Long, ByVal ad As String, ByVal kalinlik As Double, ByVal yogunluk As Double)
    gMalzemeler(idx).ad = ad
    gMalzemeler(idx).kalinlikMM = kalinlik
    gMalzemeler(idx).yogunluk = yogunluk
End Sub

Private Function JsonYoluBul() As String
    Dim kok As String, ad As String, p As String
    kok = Environ$("APPDATA") & "\Corel\"
    ad = Dir(kok & "CorelDRAW Graphics Suite *", vbDirectory)
    Do While Len(ad) > 0
        If ad <> "." And ad <> ".." Then
            p = kok & ad & "\Draw\GMS\KiloHesap\malzemeler.json"
            If Len(Dir(p)) > 0 Then
                JsonYoluBul = p
                Exit Function
            End If
            p = kok & ad & "\DrawPE\GMS\KiloHesap\malzemeler.json"
            If Len(Dir(p)) > 0 Then
                JsonYoluBul = p
                Exit Function
            End If
        End If
        ad = Dir()
    Loop
End Function

Private Sub JsonOku(ByVal yol As String)
    Dim fno As Integer, satir As String, m As Malzeme, n As Long
    Dim yeni() As Malzeme, i As Long
    On Error GoTo cik
    n = 0
    ReDim yeni(1 To 64)
    fno = FreeFile
    Open yol For Input As #fno
    Do While Not EOF(fno)
        Line Input #fno, satir
        satir = Trim$(satir)
        If InStr(1, satir, """ad""", vbTextCompare) > 0 Then
            If JsonSatirOku(satir, m) Then
                n = n + 1
                If n > UBound(yeni) Then ReDim Preserve yeni(1 To UBound(yeni) * 2)
                yeni(n) = m
            End If
        End If
    Loop
cik:
    On Error Resume Next
    Close #fno
    If n > 0 Then
        gMalzemeSayisi = n
        ReDim gMalzemeler(1 To n)
        For i = 1 To n
            gMalzemeler(i) = yeni(i)
        Next i
    End If
End Sub

Private Function JsonSatirOku(ByVal satir As String, ByRef m As Malzeme) As Boolean
    Dim ad As String, k As String, y As String
    ad = DegerAl(satir, """ad""")
    k = DegerAl(satir, """kalinlik""")
    y = DegerAl(satir, """yogunluk""")
    If Len(ad) = 0 Or Len(k) = 0 Or Len(y) = 0 Then Exit Function
    If Not IsNumeric(Val(k)) Then Exit Function
    m.ad = ad
    m.kalinlikMM = Val(k)
    m.yogunluk = Val(y)
    JsonSatirOku = (m.kalinlikMM > 0 And m.yogunluk > 0)
End Function

Private Function DegerAl(ByVal satir As String, ByVal anahtar As String) As String
    Dim p As Long, i As Long, son As Long, ch As String
    p = InStr(1, satir, anahtar, vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p, satir, ":")
    If p = 0 Then Exit Function
    i = p + 1
    Do While i <= Len(satir)
        ch = Mid$(satir, i, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        i = i + 1
    Loop
    If i > Len(satir) Then Exit Function
    If Mid$(satir, i, 1) = """" Then
        son = InStr(i + 1, satir, """")
        If son = 0 Then Exit Function
        DegerAl = Mid$(satir, i + 1, son - i - 1)
    Else
        son = i
        Do While son <= Len(satir)
            ch = Mid$(satir, son, 1)
            If ch = "," Or ch = "}" Then Exit Do
            son = son + 1
        Loop
        DegerAl = Trim$(Mid$(satir, i, son - i))
    End If
End Function

' -------------------------------------------------------------
' Hesaplama
' -------------------------------------------------------------
' Secili sekillerin toplam kilosunu hesaplar (kg).
' Donus: toplam kg. adet: sayilan sekil sayisi. hataMsj: bos = sorun yok.
Public Function SecimKiloHesapla(ByVal malzemeIdx As Long, ByRef adet As Long, ByRef hataMsj As String) As Double
    Dim toplam As Double
    adet = 0
    hataMsj = ""

    Dim doc As Document
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then
        hataMsj = "Acik dokuman yok."
        SecimKiloHesapla = 0
        Exit Function
    End If

    Dim eskiUnit As Long, eskiOpt As Boolean
    eskiUnit = doc.Unit
    eskiOpt = doc.Optimization
    On Error GoTo temiz

    doc.BeginCommandGroup "KiloHesap"
    doc.Optimization = True
    doc.Unit = cdrCentimeter

    Dim m As Malzeme
    m = gMalzemeler(malzemeIdx)

    Dim sr As ShapeRange
    Set sr = doc.Selection

    Dim i As Long, s As Shape, alan As Double
    For i = 1 To sr.Count
        Set s = sr(i)
        If s.Type <> cdrGuidelineShape Then
            alan = SekilAlanCM2(s)
            toplam = toplam + alan * (m.kalinlikMM / 10#) * m.yogunluk / 1000#
            adet = adet + 1
        End If
    Next i

temiz:
    Dim hataNo As Long, hataAciklama As String
    hataNo = Err.Number
    hataAciklama = Err.Description
    On Error Resume Next
    doc.Unit = eskiUnit
    doc.Optimization = eskiOpt
    doc.EndCommandGroup
    Application.ActiveWindow.Refresh
    If hataNo <> 0 Then
        hataMsj = "Hesap sirasinda hata: " & hataAciklama
        SecimKiloHesapla = 0
    Else
        SecimKiloHesapla = toplam
    End If
End Function

' Tek bir seklin (grup/powerclip dahil) alanini cm2 olarak dondurur.
Private Function SekilAlanCM2(ByVal s As Shape) As Double
    Dim alan As Double
    On Error GoTo cik
    Select Case s.Type
        Case cdrGroupShape
            Dim alt As Shape
            For Each alt In s.Shapes
                alan = alan + SekilAlanCM2(alt)
            Next alt
        Case cdrPowerClipShape
            Dim pc As Shape
            Set pc = Nothing
            Set pc = s.PowerClip
            If Not pc Is Nothing Then
                Dim pcAlt As Shape
                For Each pcAlt In pc.Shapes
                    alan = alan + SekilAlanCM2(pcAlt)
                Next pcAlt
            End If
        Case Else
            alan = DuzAlanCM2(s)
    End Select
cik:
    SekilAlanCM2 = alan
End Function

' Grup olmayan sekil: curve ise dogrudan, degilse gecici kopyayi
' curve'e cevirip alanini okur. Gecici kopya her durumda silinir.
Private Function DuzAlanCM2(ByVal s As Shape) As Double
    Dim alan As Double, c As Curve, tmp As Shape
    On Error GoTo cik
    If s.Type = cdrCurveShape Then
        Set c = s.Curve
        If Not c Is Nothing Then
            alan = c.Area
            GoTo cik
        End If
    End If
    Set tmp = s.Duplicate()
    tmp.ConvertToCurves
    alan = tmp.Curve.Area
    tmp.Delete
    Set tmp = Nothing
cik:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    DuzAlanCM2 = alan
End Function

' -------------------------------------------------------------
' Ozeti sayfaya yazma
' -------------------------------------------------------------
Public Sub OzetiSayfayaYaz(ByVal metin As String)
    Dim doc As Document
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then
        MsgBox "Acik dokuman yok.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If

    Dim eski As Long
    eski = doc.Unit
    On Error GoTo temiz
    doc.BeginCommandGroup "KiloHesap Ozet"
    doc.Unit = cdrCentimeter

    Dim pg As Page
    Set pg = doc.ActivePage

    Dim s As Shape
    Set s = doc.ActiveLayer.CreateArtisticText(pg.LeftEdge + 1, pg.TopEdge - 2, metin, , , "Arial", 11)

temiz:
    On Error Resume Next
    doc.Unit = eski
    doc.EndCommandGroup
    Application.ActiveWindow.Refresh
    MsgBox "Ozet cizime eklendi.", vbInformation, "Kilo Hesap"
End Sub
