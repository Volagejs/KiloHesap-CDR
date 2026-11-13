VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KiloHesapForm 
   Caption         =   "Kilo Hesap"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
End
Attribute VB_Name = "KiloHesapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------
' KiloHesap formu. Kontroller tasarim zamaninda degil,
' UserForm_Initialize icinde calisma zamaninda olusturulur.
' Boylece .frm dosyasi .frx gerektirmeden import edilebilir.
' -------------------------------------------------------------
' https://github.com/Volagejs/KiloHesap-CDR
' Lisans: MIT

Private WithEvents cmbMalzeme As MSForms.ComboBox
Private WithEvents btnHesapla As MSForms.CommandButton
Private WithEvents btnSil As MSForms.CommandButton
Private WithEvents btnSifirla As MSForms.CommandButton
Private WithEvents btnYaz As MSForms.CommandButton
Private WithEvents btnKapat As MSForms.CommandButton

Private lstSonuclar As MSForms.ListBox
Private lblToplam As MSForms.Label

Private mSatirMalzeme() As String
Private mSatirAdet() As Long
Private mSatirKg() As Double
Private mSatirSayisi As Long
Private mKapasite As Long

Private Sub UserForm_Initialize()
    mKapasite = 64
    ReDim mSatirMalzeme(1 To mKapasite)
    ReDim mSatirAdet(1 To mKapasite)
    ReDim mSatirKg(1 To mKapasite)
    mSatirSayisi = 0

    Me.Caption = "Kilo Hesap - CorelDRAW"
    Me.Width = 350
    Me.Height = 416
    Me.StartUpPosition = 1

    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblMalzeme")
    lbl.Left = 12
    lbl.Top = 8
    lbl.Width = 200
    lbl.Caption = "Malzeme / kalinlik:"

    Set cmbMalzeme = Me.Controls.Add("Forms.ComboBox.1", "cmbMalzeme")
    cmbMalzeme.Left = 12
    cmbMalzeme.Top = 24
    cmbMalzeme.Width = 190
    cmbMalzeme.Height = 18
    cmbMalzeme.Style = fmStyleDropDownList

    Set btnHesapla = Me.Controls.Add("Forms.CommandButton.1", "btnHesapla")
    btnHesapla.Left = 210
    btnHesapla.Top = 23
    btnHesapla.Width = 112
    btnHesapla.Height = 22
    btnHesapla.Caption = "Hesapla"
    btnHesapla.Default = True

    Set lstSonuclar = Me.Controls.Add("Forms.ListBox.1", "lstSonuclar")
    lstSonuclar.Left = 12
    lstSonuclar.Top = 52
    lstSonuclar.Width = 310
    lstSonuclar.Height = 214
    lstSonuclar.ColumnCount = 3
    lstSonuclar.ColumnWidths = "172 pt;38 pt;82 pt"

    Set lblToplam = Me.Controls.Add("Forms.Label.1", "lblToplam")
    lblToplam.Left = 12
    lblToplam.Top = 270
    lblToplam.Width = 310
    lblToplam.Height = 16
    lblToplam.Font.Bold = True
    lblToplam.Caption = "GENEL TOPLAM: 0.000 kg"

    Set btnSil = Me.Controls.Add("Forms.CommandButton.1", "btnSil")
    btnSil.Left = 12
    btnSil.Top = 292
    btnSil.Width = 72
    btnSil.Height = 22
    btnSil.Caption = "Satir sil"

    Set btnSifirla = Me.Controls.Add("Forms.CommandButton.1", "btnSifirla")
    btnSifirla.Left = 90
    btnSifirla.Top = 292
    btnSifirla.Width = 72
    btnSifirla.Height = 22
    btnSifirla.Caption = "Sifirla"

    Set btnYaz = Me.Controls.Add("Forms.CommandButton.1", "btnYaz")
    btnYaz.Left = 168
    btnYaz.Top = 292
    btnYaz.Width = 82
    btnYaz.Height = 22
    btnYaz.Caption = "Sayfaya yaz"

    Set btnKapat = Me.Controls.Add("Forms.CommandButton.1", "btnKapat")
    btnKapat.Left = 256
    btnKapat.Top = 292
    btnKapat.Width = 66
    btnKapat.Height = 22
    btnKapat.Caption = "Kapat"

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblIpucu")
    lbl.Left = 12
    lbl.Top = 322
    lbl.Width = 320
    lbl.Height = 24
    lbl.Font.Size = 8
    lbl.Caption = "Sekilleri CorelDRAW penceresinde secin, malzemeyi secip Hesapla'ya basin. Pencere acik kalabilir, her hesap listeye eklenir."

    MalzemeleriDoldur
End Sub

Private Sub MalzemeleriDoldur()
    Dim i As Long
    cmbMalzeme.Clear
    For i = 1 To gMalzemeSayisi
        cmbMalzeme.AddItem gMalzemeler(i).ad & " " & KalinlikYaz(gMalzemeler(i).kalinlikMM) & " mm"
    Next i
    If gMalzemeSayisi > 0 Then cmbMalzeme.ListIndex = 0
End Sub

Private Function KalinlikYaz(ByVal k As Double) As String
    If k = Int(k) Then
        KalinlikYaz = Format$(k, "0")
    Else
        KalinlikYaz = Format$(k, "0.00")
    End If
End Function

' -------------------------------------------------------------
' Olaylar
' -------------------------------------------------------------
Private Sub btnHesapla_Click()
    If gMalzemeSayisi < 1 Then
        MsgBox "Malzeme listesi yuklenemedi.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If
    If cmbMalzeme.ListIndex < 0 Then
        MsgBox "Once malzeme/kalinlik secin.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If

    Dim idx As Long, adet As Long, hata As String, kg As Double
    idx = cmbMalzeme.ListIndex + 1
    kg = SecimKiloHesapla(idx, adet, hata)

    If Len(hata) > 0 Then
        MsgBox hata, vbExclamation, "Kilo Hesap"
        Exit Sub
    End If
    If adet = 0 Then
        MsgBox "Secili sekil bulunamadi. Once CorelDRAW'da sekil secin.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If

    SatirEkle cmbMalzeme.List(cmbMalzeme.ListIndex), adet, kg
End Sub

Private Sub btnSil_Click()
    Dim i As Long, j As Long
    i = lstSonuclar.ListIndex
    If i < 0 Then
        MsgBox "Once listeden bir satir secin.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If
    For j = i + 2 To mSatirSayisi
        mSatirMalzeme(j - 1) = mSatirMalzeme(j)
        mSatirAdet(j - 1) = mSatirAdet(j)
        mSatirKg(j - 1) = mSatirKg(j)
    Next j
    mSatirSayisi = mSatirSayisi - 1
    lstSonuclar.RemoveItem i
    ToplamGuncelle
End Sub

Private Sub btnSifirla_Click()
    If mSatirSayisi = 0 Then Exit Sub
    If MsgBox("Tum sonuc listesi temizlenecek. Emin misiniz?", vbYesNo + vbQuestion, "Kilo Hesap") <> vbYes Then Exit Sub
    mSatirSayisi = 0
    lstSonuclar.Clear
    ToplamGuncelle
End Sub

Private Sub btnYaz_Click()
    If mSatirSayisi = 0 Then
        MsgBox "Liste bos. Once hesaplama yapin.", vbExclamation, "Kilo Hesap"
        Exit Sub
    End If
    OzetiSayfayaYaz OzetMetni
End Sub

Private Sub btnKapat_Click()
    Unload Me
End Sub

' -------------------------------------------------------------
' Liste / toplam yonetimi
' -------------------------------------------------------------
Private Sub SatirEkle(ByVal malzemeAd As String, ByVal adet As Long, ByVal kg As Double)
    If mSatirSayisi >= mKapasite Then
        mKapasite = mKapasite * 2
        ReDim Preserve mSatirMalzeme(1 To mKapasite)
        ReDim Preserve mSatirAdet(1 To mKapasite)
        ReDim Preserve mSatirKg(1 To mKapasite)
    End If
    mSatirSayisi = mSatirSayisi + 1
    mSatirMalzeme(mSatirSayisi) = malzemeAd
    mSatirAdet(mSatirSayisi) = adet
    mSatirKg(mSatirSayisi) = kg

    With lstSonuclar
        .AddItem malzemeAd
        .List(.ListCount - 1, 1) = CStr(adet)
        .List(.ListCount - 1, 2) = Format$(kg, "0.000")
    End With
    ToplamGuncelle
End Sub

Private Sub ToplamGuncelle()
    Dim t As Double, i As Long
    For i = 1 To mSatirSayisi
        t = t + mSatirKg(i)
    Next i
    lblToplam.Caption = "GENEL TOPLAM: " & Format$(t, "0.000") & " kg"
End Sub

' Ayni malzemenin satirlarini birlestirip ozet metni uretir.
Private Function OzetMetni() As String
    Dim adlar() As String, adetler() As Long, kgler() As Double
    Dim n As Long, i As Long, j As Long, bulundu As Long

    If mSatirSayisi < 1 Then Exit Function
    ReDim adlar(1 To mSatirSayisi)
    ReDim adetler(1 To mSatirSayisi)
    ReDim kgler(1 To mSatirSayisi)

    For i = 1 To mSatirSayisi
        bulundu = 0
        For j = 1 To n
            If StrComp(adlar(j), mSatirMalzeme(i), vbTextCompare) = 0 Then
                bulundu = j
                Exit For
            End If
        Next j
        If bulundu = 0 Then
            n = n + 1
            adlar(n) = mSatirMalzeme(i)
            adetler(n) = mSatirAdet(i)
            kgler(n) = mSatirKg(i)
        Else
            adetler(bulundu) = adetler(bulundu) + mSatirAdet(i)
            kgler(bulundu) = kgler(bulundu) + mSatirKg(i)
        End If
    Next i

    Dim t As Double, metin As String
    metin = "KILO HESAP OZETI" & vbCrLf & Format$(Now, "dd.mm.yyyy hh:nn") & vbCrLf & String$(34, "-") & vbCrLf
    For i = 1 To n
        metin = metin & adlar(i) & ": " & adetler(i) & " adet  -  " & Format$(kgler(i), "0.000") & " kg" & vbCrLf
        t = t + kgler(i)
    Next i
    metin = metin & String$(34, "-") & vbCrLf & "GENEL TOPLAM: " & Format$(t, "0.000") & " kg"
    OzetMetni = metin
End Function
