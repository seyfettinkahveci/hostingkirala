VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form KullaniciBilgileri 
   BackColor       =   &H8000000D&
   Caption         =   "KullaniciBilgileri"
   ClientHeight    =   8820
   ClientLeft      =   2985
   ClientTop       =   1485
   ClientWidth     =   13185
   LinkTopic       =   "Form2"
   ScaleHeight     =   8820
   ScaleWidth      =   13185
   Begin VB.Frame UyeKaydiFrame 
      BackColor       =   &H8000000D&
      Caption         =   "Üye Ekle"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   7215
      Begin VB.ComboBox Yetki 
         Height          =   315
         ItemData        =   "KullaniciBilgileri.frx":0000
         Left            =   3000
         List            =   "KullaniciBilgileri.frx":000A
         TabIndex        =   20
         Text            =   "Seçiniz"
         Top             =   4680
         Width           =   3015
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
         TabIndex        =   11
         Top             =   2400
         Width           =   3000
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
         TabIndex        =   10
         Top             =   3000
         Width           =   3000
      End
      Begin VB.CommandButton UyaOlIptal 
         Caption         =   "Geri Dön"
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
         TabIndex        =   9
         Top             =   5520
         Width           =   2000
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
         TabIndex        =   8
         Top             =   3600
         Width           =   1500
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
         TabIndex        =   7
         Top             =   3600
         Width           =   1140
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
         TabIndex        =   5
         Top             =   1200
         Width           =   3000
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
         TabIndex        =   4
         Top             =   600
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
         Left            =   1320
         TabIndex        =   3
         Top             =   5520
         Width           =   2000
      End
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
         TabIndex        =   2
         Top             =   4080
         Width           =   3000
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "Yetki"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   4680
         Width           =   1335
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
         TabIndex        =   18
         Top             =   2400
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
         TabIndex        =   17
         Top             =   3000
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
         TabIndex        =   16
         Top             =   3600
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
         TabIndex        =   15
         Top             =   1800
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
         TabIndex        =   13
         Top             =   600
         Width           =   1500
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
         TabIndex        =   12
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12240
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"KullaniciBilgileri.frx":0020
      OLEDBString     =   $"KullaniciBilgileri.frx":00FC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select KullaniciAdi, AdSoyad, Tel, Cinsiyet, Adres,Sifre from Kullanicilar "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "KullaniciBilgileri.frx":01D8
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kullanýcý Bilgileri"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "KullaniciAdi"
         Caption         =   "KullaniciAdi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "AdSoyad"
         Caption         =   "AdSoyad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Tel"
         Caption         =   "Tel"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Cinsiyet"
         Caption         =   "Cinsiyet"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Adres"
         Caption         =   "Adres"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Sifre"
         Caption         =   "Sifre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "KullaniciBilgileri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KullaniciA As String
Public Adres As String

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub

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
Unload Me
Yonetici.Show
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
ElseIf Yetki.Text = "Seçiniz" Then
dugme = MsgBox("Lütfen Yetki Seçiniz", 64, "Uyari")
Else
    If Yetki.Text = "Kullanici" Then
    Yetki = 1
    Else
    Yetki = 2
    End If
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kullanicilar")
Rs.AddNew
Rs!AdSoyad = AdSoyadText.Text
Rs!KullaniciAdi = KullaniciAdiText.Text
Rs!Sifre = SifreText.Text
Rs!Yetki = Yetki
If (Option2.Value = True) Then
Cinsiyet = "E"
End If
If (Option1.Value = True) Then
Cinsiyet = "K"
End If
Rs!Cinsiyet = Cinsiyet
Rs!Adres = AdresText.Text
Rs!Tel = TelText.Text
Rs.Update
Db.Close
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub
