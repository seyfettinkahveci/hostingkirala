VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PaketEkleGuncelle 
   BackColor       =   &H8000000D&
   Caption         =   "Paket Ekle Guncelle"
   ClientHeight    =   2805
   ClientLeft      =   4020
   ClientTop       =   4980
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   ScaleHeight     =   2805
   ScaleWidth      =   10650
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Yeni Paket Ekleme "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.TextBox OzellikText 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   540
         Width           =   1935
      End
      Begin VB.ComboBox Kategori 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Text            =   "Seçiniz"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Geri Dön"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ekle"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox FiyatText 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox PaketKoduText 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Özellikleri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   550
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Kategori"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Fiyat"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Alan Kodu"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PaketEkleGuncelle.frx":0000
      Height          =   2775
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4895
      _Version        =   393216
      BackColor       =   -2147483637
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sistemde Bulunan Paketler"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "PaketKodu"
         Caption         =   "PaketKodu"
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
         DataField       =   "Fiyat"
         Caption         =   "Fiyat"
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
         DataField       =   "Ozellik"
         Caption         =   "Ozellik"
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
         DataField       =   "Kategori"
         Caption         =   "Kategori"
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
         BeginProperty Column00 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   120
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"PaketEkleGuncelle.frx":0015
      OLEDBString     =   $"PaketEkleGuncelle.frx":00F1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select PaketKodu,Fiyat,Ozellik,Kategori  from HostingBilgi"
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
End
Attribute VB_Name = "PaketEkleGuncelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
If PaketKoduText.Text = "" Then
dugme = MsgBox("Alan Kodu Boþ Olamaz", 64, "Uyari")
ElseIf FiyatText.Text = "" Then
dugme = MsgBox("Fiyat Boþ Olamaz", 64, "Uyari")
ElseIf Kategori.Text = "Seçiniz" Then
dugme = MsgBox("Kategori Seçiniz", 64, "Uyari")
ElseIf OzellikText.Text = "Seçiniz" Then
dugme = MsgBox("Özellik Seçiniz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("HostingBilgi")
Rs.AddNew
Rs!PaketKodu = PaketKoduText.Text
Rs!Fiyat = FiyatText.Text
Rs!Ozellik = OzellikText.Text
Rs!Kategori = Kategori.Text
Rs.Update
Db.Close
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Command3_Click()
Unload Me
Yonetici.Show
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub


Private Sub FiyatText_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Form_Load()
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kategoriler")
Do Until Rs.EOF
KategoriAdi = Rs("KategoriAdi")
Kategori.AddItem KategoriAdi
Rs.MoveNext
Loop

Db.Close
End Sub
