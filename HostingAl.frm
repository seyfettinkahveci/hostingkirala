VERSION 5.00
Begin VB.Form HostingAl 
   BackColor       =   &H8000000D&
   Caption         =   "HostingAl"
   ClientHeight    =   4800
   ClientLeft      =   4635
   ClientTop       =   4980
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   9525
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
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sepete Ekle"
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
      Left            =   6600
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Height          =   4155
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Paket Özellikleri"
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
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Paket Seçiniz"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Kategori Seçiniz"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "HostingAl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KullaniciA As String
Private Sub Command1_Click()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Sepet")
Rs.AddNew
Rs!KullaniciAdi = UyeGirisi.KullaniciA
Rs!PaketKodu = List2.Text
Rs.Update
Db.Close
End Sub

Private Sub Command2_Click()
Unload Me
Kullanici.Show
End Sub

Private Sub Form_Load()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kategoriler")
Do Until Rs.EOF
KategoriAdi = Rs("KategoriAdi")
List1.AddItem KategoriAdi
Rs.MoveNext
Loop
Db.Close
End Sub

Private Sub List1_Click()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("HostingBilgi")
List2.Clear
Do Until Rs.EOF
KategoriAdi = Rs("Kategori")
PaketKodu = Rs("PaketKodu")
If KategoriAdi = List1.Text Then
List2.AddItem PaketKodu
End If
Rs.MoveNext
Loop
List2.Visible = True
Label2.Visible = True
Db.Close
End Sub

Private Sub List2_Click()
Label3.Visible = True
Text1.Visible = True
Command1.Visible = True
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("HostingBilgi")
Text1.Text = ""
Do Until Rs.EOF
PaketKodu = Rs("PaketKodu")
Ozellik = Rs("Ozellik")
If PaketKodu = List2.Text Then
Text1.Text = Ozellik
End If
Rs.MoveNext
Loop
List2.Visible = True
Label2.Visible = True
Db.Close
End Sub
