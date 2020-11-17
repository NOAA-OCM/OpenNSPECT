VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "N-SPECTor Gadget"
   ClientHeight    =   3435
   ClientLeft      =   3930
   ClientTop       =   2700
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2655
      Top             =   2280
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   270
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2565
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Coding support for this project made possible by N-SPECTor Gadget!"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   360
      TabIndex        =   1
      Top             =   2730
      Width           =   1905
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Timer1.Interval = 4000
    
End Sub

Private Sub Timer1_Timer()
    Unload frmAbout
End Sub
