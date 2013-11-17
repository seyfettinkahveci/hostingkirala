VERSION 5.00
Begin VB.Form UyeKayit 
   BackColor       =   &H8000000D&
   Caption         =   "Kayýt Ol"
   ClientHeight    =   6780
   ClientLeft      =   5655
   ClientTop       =   2505
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   7470
   Begin VB.Frame UyeKaydiFrame 
      BackColor       =   &H8000000D&
      Caption         =   "Üye Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox TelText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   17
         Top             =   4080
         Width           =   3000
      End
      Begin VB.CommandButton UyeOlButon 
         Caption         =   "Üye Ol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1440
         TabIndex        =   9
         Top             =   4800
         Width           =   2000
      End
      Begin VB.TextBox AdSoyadText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   3000
      End
      Begin VB.TextBox KullaniciAdiText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   7
         Top             =   1200
         Width           =   3000
      End
      Begin VB.TextBox SifreText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   3000
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000D&
         Caption         =   "Kýz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   400
         Left            =   3000
         TabIndex        =   5
         Top             =   3600
         Width           =   1140
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000D&
         Caption         =   "Erkek"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   400
         Left            =   4320
         TabIndex        =   4
         Top             =   3600
         Width           =   1500
      End
      Begin VB.CommandButton UyaOlIptal 
         Caption         =   "Ýptal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3960
         TabIndex        =   3
         Top             =   4800
         Width           =   2000
      End
      Begin VB.TextBox AdresText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   2
         Top             =   3000
         Width           =   3000
      End
      Begin VB.TextBox SifreTekrarText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2400
         Width           =   3000
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Tel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1500
         TabIndex        =   16
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Ad Soyad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   400
         Index           =   1
         Left            =   1500
         TabIndex        =   15
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Kullanýcý Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Index           =   1
         Left            =   1500
         TabIndex        =   14
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Index           =   1
         Left            =   1500
         TabIndex        =   13
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Cinsiyet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Index           =   1
         Left            =   1500
         TabIndex        =   12
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label label5 
         BackColor       =   &H8000000D&
         Caption         =   "Adres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1500
         TabIndex        =   11
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Þifre Tekrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1500
         TabIndex        =   10
         Top             =   2400
         Width           =   1500
      End
   End
End
Attribute VB_Name = "UyeKayit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database
Dim Rs As Recordset

Private Sub Option1_Click()
Image1.Visible = True
Image1.Picture = LoadPicture("Resimler/Woman.jpg")
End Sub

Private Sub Option2_Click()
Image1.Visible = True
Image1.Picture = LoadPicture("Resimler/Man.jpg")
End Sub
Private Sub TelText_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TelText_LostFocus()
TelText.Text = Format(TelText.Text, "0(###) ### ## ##")
End Sub

Private Sub UyaOlIptal_Click()
Form1.Show
UyeKayit.Hide

End Sub



Private Sub UyeOlButon_Click()
If AdSoyadText.Text = "" Then
dugme = MsgBox("Ad Soyad Boþ Olamaz", 64, "Uyari")
ElseIf KullaniciAdiText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf SifreText.Text <> SifreTekrarText.Text Then
dugme = MsgBox("Þifreler Eþleþmiyor", 64, "Uyari")
ElseIf SifreText.Text = "" Or SifreTekrarText.Text = "" Then
dugme = MsgBox("Þifreler Boþ Olamaz", 64, "Uyari")
ElseIf AdresText.Text = "" Then
dugme = MsgBox("Adres Boþ olamaz", 64, "Uyari")
ElseIf Option2.Value = False And Option1.Value = False Then
dugme = MsgBox("Cinsiyet Seçiniz", 64, "Uyari")
ElseIf TelText.Text = "" Then
dugme = MsgBox("Tel Boþ Olamaz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kullanicilar")
Rs.AddNew
Rs!AdSoyad = AdSoyadText.Text
Rs!KullaniciAdi = KullaniciAdiText.Text
Rs!Sifre = SifreText.Text
If (Option2.Value = True) Then
Cinsiyet = "E"
End If
If (Option1.Value = True) Then
Cinsiyet = "K"
End If
Rs!Cinsiyet = Cinsiyet
Rs!Adres = AdresText.Text
Rs!Tel = TelText.Text
Rs!Yetki = 1
Rs.Update
Db.Close
UyeKayit.Hide
UyeGirisi.Show
End If
End Sub

