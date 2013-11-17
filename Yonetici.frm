VERSION 5.00
Begin VB.Form Yonetici 
   BackColor       =   &H8000000D&
   Caption         =   "Yonetici"
   ClientHeight    =   3240
   ClientLeft      =   5445
   ClientTop       =   4365
   ClientWidth     =   8415
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   8415
   Begin VB.CommandButton Command6 
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
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tüm Hostingler "
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
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   2500
   End
   Begin VB.CommandButton Command4 
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
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Onay Bekleyenler"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paket Ekle/Güncelle"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kategori Ekle/Güncelle"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2500
   End
   Begin VB.Label KazanilanParaBilgi 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Yonetici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
KategoriEkleGuncelle.Show
End Sub

Private Sub Command2_Click()
PaketEkleGuncelle.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
OnayBekleyenler.Show
End Sub

Private Sub Command4_Click(Index As Integer)
KullaniciBilgileri.Show
Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
TumHostingler.Show
End Sub

Private Sub Command6_Click()
 End
End Sub

Private Sub Form_Load()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("HostingAlan")
Fiyat = 0
Toplam = 0
Do Until Rs.EOF
If Rs("Onay") = 2 Then
Fiyat = Fiyat + Rs("Fiyat")
End If
If Rs("Onay") = 1 Then
Toplam = Toplam + 1
End If
Rs.MoveNext
Loop
KazanilanParaBilgi.Caption = " Þuana Kadar Toplam " & Fiyat & " TL Kazandýnýz"
Command3.Caption = " Onay Bekleyenler(" & Toplam & ")"
Db.Close
End Sub
