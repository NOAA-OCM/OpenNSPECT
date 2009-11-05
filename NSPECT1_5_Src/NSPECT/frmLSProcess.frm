VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmLSProcess 
   Caption         =   "Processing Length Slope Grid"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   735
      TabIndex        =   1
      Top             =   660
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   120
   End
   Begin VB.Label lbl1 
      Caption         =   "Processing Length/Slope Grid..."
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   210
      Width           =   2970
   End
End
Attribute VB_Name = "frmLSProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    With Timer1
         .Enabled = True
         .Interval = 3000
    End With
    
    DoEvents
    
End Sub



Private Sub Timer1_Timer()
    
    Dim intValue As Integer
    intValue = Int((100 * Rnd) + 1)
    
    ProgressBar1.Value = intValue
    
End Sub
