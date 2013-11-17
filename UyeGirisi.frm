VERSION 5.00
Begin VB.Form UyeGirisi 
   BackColor       =   &H8000000D&
   Caption         =   "Üye Giriþi"
   ClientHeight    =   5565
   ClientLeft      =   7500
   ClientTop       =   3945
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   ScaleHeight     =   5565
   ScaleWidth      =   4200
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Kullanici Bilgileriniz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Parolamý Unuttum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Giriþ Yap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox SifreText 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox KAdiText 
         Height          =   350
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Parola"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Kullanýcý Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   $"UyeGirisi.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   4095
   End
End
Attribute VB_Name = "UyeGirisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database
Dim Rs As Recordset
Dim Mantik As Boolean
Public KullaniciA As String
Public Adres As String
Private Sub Command1_Click()
Mantik = True
If KAdiText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf SifreText.Text = "" Then
dugme = MsgBox("Þifre Boþ Olamaz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kullanicilar")
Do Until Rs.EOF
KullaniciAdi = Rs("KullaniciAdi")
Sifre = Rs("Sifre")
Yetki = Rs("Yetki")
If KullaniciAdi = KAdiText.Text And Sifre = SifreText.Text And Yetki = 2 Then
KullaniciA = KullaniciAdi
Yonetici.Show
Unload Me
Mantik = False
ElseIf KullaniciAdi = KAdiText.Text And Sifre = SifreText.Text And Yetki = 1 Then
KullaniciA = KullaniciAdi
Kullanici.Show
Unload Me
Mantik = False
End If
Rs.MoveNext
Loop
Db.Close
If Mantik = True Then
dugme = MsgBox("Yanlýþ Kullanýcý Adý Parola Lütfen Tekrar Deneyin", 64, "Uyari")
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command3_Click()
Unload Me
ParolamiAnimsa.Show
End Sub
