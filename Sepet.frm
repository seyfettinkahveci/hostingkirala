VERSION 5.00
Begin VB.Form Sepet 
   BackColor       =   &H8000000D&
   Caption         =   "Sepet"
   ClientHeight    =   4290
   ClientLeft      =   6675
   ClientTop       =   4770
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   6075
   Begin VB.Frame Frame1 
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Alýþveriþi Tamamla"
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
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Seçiliyi Sepetten Kaldýr"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ListBox AlanAdi 
         Height          =   2010
         Left            =   3000
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.ListBox Paket 
         Height          =   2010
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Alan Adý Bilgileri"
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
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Paket Bilgileri"
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
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Sepet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Kullanici.Show
End Sub

Private Sub Command2_Click()
If Paket.Text <> "" Then
     Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    SQL = "Delete From Sepet Where PaketKodu='" & Paket.Text & "'"
    Db.Execute (SQL)
    Db.Close
    Paket.Clear
    AlanAdi.Clear
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    Set Rs = Db.OpenRecordset("Sepet")
    Do Until Rs.EOF
    KullaniciAdi = Rs("KullaniciAdi")
    If (KullaniciAdi = UyeGirisi.KullaniciA) Then
    If Rs("PaketKodu") <> "" Then
    Paket.AddItem Rs("PaketKodu")
    End If
    If Rs("AlanAdi") <> "" Then
    AlanAdi.AddItem Rs("AlanAdi")
    End If
    
    End If
    Rs.MoveNext
    Loop
    Db.Close
End If
If AlanAdi.Text <> "" Then
     Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    SQL2 = "Delete From Sepet Where AlanAdi='" & AlanAdi.Text & "'"
    Db.Execute (SQL2)
    Db.Close
    Paket.Clear
    AlanAdi.Clear
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    Set Rs = Db.OpenRecordset("Sepet")
    Do Until Rs.EOF
    KullaniciAdi = Rs("KullaniciAdi")
    If (KullaniciAdi = UyeGirisi.KullaniciA) Then
    If Rs("PaketKodu") <> "" Then
    Paket.AddItem Rs("PaketKodu")
    End If
    If Rs("AlanAdi") <> "" Then
    AlanAdi.AddItem Rs("AlanAdi")
    End If
    
    End If
    Rs.MoveNext
    Loop
    Db.Close
End If
End Sub

Private Sub Command3_Click()
If Paket.Text = "" Or AlanAdi.Text = "" Then
    dugme = MsgBox("Lütfen Sepetinizde Almak Ýstediðiniz Alan Adý Ýle Paketinizi Birlikte Seçiniz", 64, "Uyari")
Else
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    Set Rs = Db.OpenRecordset("HostingBilgi")
    Do Until Rs.EOF
    PaketKodu = Rs("PaketKodu")
    If (PaketKodu = Paket.Text) Then
        Fiyat = Rs("Fiyat") + 20
    End If
    Rs.MoveNext
    Loop
    Db.Close
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    Set Rs = Db.OpenRecordset("HostingAlan")
    Rs.AddNew
    Rs!AlanAdi = AlanAdi.Text
    Rs!PaketKodu = Paket.Text
    Rs!KullaniciAdi = UyeGirisi.KullaniciA
    Rs!AlisTarihi = Day(Now) & "." & Month(Now) & "." & Year(Now)
    Rs!SonTarih = Day(Now) & "." & Month(Now) & "." & (Year(Now) + 1)
    Rs!Fiyat = Fiyat
    Rs!Onay = 1
    Rs.Update
    Db.Close
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    SQL = "Delete From Sepet Where PaketKodu='" & Paket.Text & "'"
    SQL2 = "Delete From Sepet Where AlanAdi='" & AlanAdi.Text & "'"
    Db.Execute (SQL)
    Db.Execute (SQL2)
    Db.Close
    Paket.Clear
    AlanAdi.Clear
    Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
    Set Rs = Db.OpenRecordset("Sepet")
    Do Until Rs.EOF
    KullaniciAdi = Rs("KullaniciAdi")
    If (KullaniciAdi = UyeGirisi.KullaniciA) Then
    If Rs("PaketKodu") <> "" Then
    Paket.AddItem Rs("PaketKodu")
    End If
    If Rs("AlanAdi") <> "" Then
    AlanAdi.AddItem Rs("AlanAdi")
    End If
    
    End If
    Rs.MoveNext
    Loop
    Db.Close
End If
End Sub

Private Sub Form_Load()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Sepet")
Do Until Rs.EOF
KullaniciAdi = Rs("KullaniciAdi")
If (KullaniciAdi = UyeGirisi.KullaniciA) Then
    If Rs("PaketKodu") <> "" Then
    Paket.AddItem Rs("PaketKodu")
    End If
    If Rs("AlanAdi") <> "" Then
    AlanAdi.AddItem Rs("AlanAdi")
    End If
    
End If
Rs.MoveNext
Loop
Db.Close
End Sub
