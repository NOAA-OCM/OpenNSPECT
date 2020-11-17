VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProgress 
   Caption         =   "Processing..."
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1290
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   4365
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   315
         Left            =   390
         TabIndex        =   1
         Top             =   810
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   405
         TabIndex        =   2
         Top             =   165
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event HasStopped()



Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent HasStopped
End Sub

