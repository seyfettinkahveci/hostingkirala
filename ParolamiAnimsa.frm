VERSION 5.00
Begin VB.Form ParolamiAnimsa 
   BackColor       =   &H8000000D&
   Caption         =   "Parolami Animsa"
   ClientHeight    =   4110
   ClientLeft      =   5850
   ClientTop       =   4770
   ClientWidth     =   7110
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   7110
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Kullanýcý Bilgileriniz"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "Parolamý Deðiþtir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox KullaniciText 
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
         Left            =   2835
         TabIndex        =   0
         Top             =   480
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
         Left            =   2835
         TabIndex        =   1
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox YeniSifreText 
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
         Left            =   2835
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   3000
      End
      Begin VB.TextBox YeniSifreT 
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
         Left            =   2835
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2280
         Width           =   3000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Geri Dön"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "KullaniciAdi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   795
         TabIndex        =   10
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Ad Soyad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Yeni Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   795
         TabIndex        =   8
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Yeni Þifre Tekrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   795
         TabIndex        =   7
         Top             =   2280
         Width           =   1995
      End
   End
End
Attribute VB_Name = "ParolamiAnimsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If KullaniciText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf AdSoyadText.Text = "" Then
dugme = MsgBox("Ad Soyadý Boþ Býrakmayýn", 64, "Uyari")
ElseIf YeniSifreText.Text = "" Or YeniSifreT.Text = "" Then
dugme = MsgBox("Þifreler Boþ Olamaz", 64, "Uyari")
ElseIf YeniSifreText.Text <> YeniSifreT.Text Then
dugme = MsgBox("Þifreler Eþleþmemektedir", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
SQL = "UPDATE Kullanicilar SET Sifre='" & YeniSifreText.Text & "' WHERE KullaniciAdi='" & KullaniciText.Text & "' and AdSoyad='" & AdSoyadText.Text & "'"
Db.Execute (SQL)
Db.Close
Unload Me
UyeGirisi.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
UyeGirisi.Show
End Sub
