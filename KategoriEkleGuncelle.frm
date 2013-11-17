VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form KategoriEkleGuncelle 
   BackColor       =   &H8000000D&
   Caption         =   "KategoriEkleGuncelle"
   ClientHeight    =   2265
   ClientLeft      =   5655
   ClientTop       =   4980
   ClientWidth     =   7455
   LinkTopic       =   "Form2"
   ScaleHeight     =   2265
   ScaleWidth      =   7455
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Yeni Kategori Ekle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   4215
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
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ekle"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox KategoriAdiText 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   855
         Left            =   0
         Top             =   1920
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
         Connect         =   $"KategoriEkleGuncelle.frx":0000
         OLEDBString     =   $"KategoriEkleGuncelle.frx":00DC
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select KategoriAdi from Kategoriler"
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
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Kategori Adý"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "KategoriEkleGuncelle.frx":01B8
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3201
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
      Caption         =   "Sistemde Bulunan Kategoriler"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "KategoriAdi"
         Caption         =   "KategoriAdi"
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
      EndProperty
   End
End
Attribute VB_Name = "KategoriEkleGuncelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If KategoriAdiText = "" Then
dugme = MsgBox("Kategori Adý Boþ Olamaz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kategoriler")
Rs.AddNew
Rs!KategoriAdi = KategoriAdiText.Text
Rs.Update
Db.Close
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Command2_Click()
Unload Me
Yonetici.Show
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub
