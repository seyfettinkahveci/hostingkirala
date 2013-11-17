VERSION 5.00
Begin VB.Form AlanAdiAl 
   BackColor       =   &H8000000D&
   Caption         =   "AlanAdiAl"
   ClientHeight    =   3945
   ClientLeft      =   6675
   ClientTop       =   4980
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3945
   ScaleWidth      =   4860
   Begin VB.CommandButton Command3 
      Caption         =   "Sepete Ekle"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4575
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
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Alan Adý Sorgulama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Sorgula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Alan Adý Giriniz:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Bilgi 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "AlanAdiAl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("HostingAlan")
Do Until Rs.EOF
AlanAdi = Rs("AlanAdi")
    If AlanAdi = Text1.Text Then
        Bool = True
    End If
Rs.MoveNext
Loop
    If Bool = True Then
        Bilgi.Caption = Text1.Text + " Alan Adý Daha Önceden Alýnmýþtýr"
    Else
        Bilgi.Caption = Text1.Text + " Alan Adý Satýn Almaya Müsaittir"
        Command3.Enabled = True
    End If
Db.Close
End Sub

Private Sub Command2_Click()
Unload Me
Kullanici.Show
End Sub

Private Sub Command3_Click()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Sepet")
Rs.AddNew
Rs!KullaniciAdi = UyeGirisi.KullaniciA
Rs!AlanAdi = Text1.Text
Rs.Update
Db.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
   dugme = MsgBox("Alan Adýnda Boþluk Olamaz", 64, "Uyari")
   KeyAscii = 0
End If
End Sub
