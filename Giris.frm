VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kahveci Hosting Sistemine Hoþgeldiniz"
   ClientHeight    =   5685
   ClientLeft      =   5250
   ClientTop       =   4980
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8850
   Begin VB.Image Image4 
      Height          =   570
      Left            =   3120
      MousePointer    =   1  'Arrow
      Picture         =   "Giris.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   3120
      MousePointer    =   1  'Arrow
      Picture         =   "Giris.frx":167A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   3120
      Picture         =   "Giris.frx":3030
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   -360
      Picture         =   "Giris.frx":558D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Unload Me
UyeGirisi.Show
End Sub

Private Sub Image3_Click()
Form1.Hide
UyeKayit.Show

End Sub

Private Sub Image4_Click()
End
End Sub

Private Sub Label1_Click()

End Sub
