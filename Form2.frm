VERSION 5.00
Begin VB.Form Kullanici 
   BackColor       =   &H8000000D&
   Caption         =   "Kullanici"
   ClientHeight    =   2910
   ClientLeft      =   5850
   ClientTop       =   4980
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   7980
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "Sepetim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "Çýkýþ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5280
      TabIndex        =   5
      Top             =   1440
      Width           =   2500
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "Kullanýcý Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "Hosting Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Alan Adý Al"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Hosting Al"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2500
   End
End
Attribute VB_Name = "Kullanici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
HostingAl.Show
End Sub

Private Sub Command2_Click()
Unload Me
AlanAdiAl.Show
End Sub

Private Sub Command3_Click()
Unload Me
HostingBilgileri.Show
End Sub

Private Sub Command4_Click()
KullaniciBilgileri2.Show
Unload Me
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Unload Me
Sepet.Show
End Sub

Private Sub Form_Load()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Sepet")
Toplam = 0
Do Until Rs.EOF
If Rs("KullaniciAdi") = UyeGirisi.KullaniciA Then
Toplam = Toplam + 1
End If
Rs.MoveNext
Loop
Command6.Caption = " Sepetim(" & Toplam & ")"
Db.Close
End Sub
