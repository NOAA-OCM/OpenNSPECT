VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAddPoll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Pollutant"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTabStripContainer 
      BorderStyle     =   0  'None
      Caption         =   "Coefficients"
      Height          =   4600
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   819
      Width           =   7384
      Begin VB.ComboBox cboCoeffSet 
         Height          =   315
         ItemData        =   "frmAddPoll.frx":0000
         Left            =   4914
         List            =   "frmAddPoll.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   75
         Width           =   2200
      End
      Begin VB.TextBox txtCoeffSet 
         Height          =   285
         Left            =   1305
         TabIndex        =   18
         Top             =   60
         Width           =   2000
      End
      Begin VB.TextBox txtCoeffSetDesc 
         Height          =   285
         Left            =   1290
         TabIndex        =   5
         Top             =   468
         Width           =   5820
      End
      Begin VB.TextBox txtActiveCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4125
         TabIndex        =   6
         Top             =   3585
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid grdPollDef 
         Height          =   3495
         Left            =   180
         TabIndex        =   7
         Top             =   1200
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   7
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   0
         FillStyle       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Coefficient Set"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   7155
         TabIndex        =   13
         Top             =   915
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   12
         Top             =   915
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coefficients"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3945
         TabIndex        =   11
         Top             =   915
         Width           =   3210
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   10
         Top             =   915
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   9
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Land Cover Type"
         Height          =   260
         Index           =   7
         Left            =   3510
         TabIndex        =   8
         Top             =   117
         Width           =   1235
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6885
      TabIndex        =   2
      Top             =   5895
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5775
      TabIndex        =   1
      Top             =   5895
      Width           =   975
   End
   Begin VB.TextBox txtPollName 
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Top             =   90
      Width           =   2000
   End
   Begin VB.Frame fraTabStripContainer 
      BorderStyle     =   0  'None
      Caption         =   "Coefficients"
      Height          =   4600
      Index           =   2
      Left            =   150
      TabIndex        =   16
      Top             =   819
      Width           =   7384
      Begin MSFlexGridLib.MSFlexGrid grdWQStd 
         Height          =   4004
         Left            =   0
         TabIndex        =   17
         Top             =   117
         Width           =   7202
         _ExtentX        =   12700
         _ExtentY        =   7064
         _Version        =   393216
         Cols            =   4
         ScrollBars      =   2
      End
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   5235
      Left            =   180
      TabIndex        =   15
      Top             =   450
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   9234
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Coefficients"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Water Quality Standards"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Pollutant Name:"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   105
      Width           =   1785
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuCopyCoeffSet 
         Caption         =   "&Copy Coefficient Set..."
      End
      Begin VB.Menu mnuImportCoeffSet 
         Caption         =   "&Import Coefficient Set"
      End
   End
End
Attribute VB_Name = "frmAddPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mintCurFrame As Integer ' Current Frame visible in Tabstrip



Private Sub Form_Load()
   mintCurFrame = 1
   InitPollDefGrid grdPollDef
   InitPollWQStdGrid grdWQStd

End Sub

Private Sub cmdCancel_Click()

   Unload Me

End Sub

Private Sub mnuCopyCoeffSet_Click()

    frmCopyCoeffSet.Show vbModeless, Me

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Control tabstrip functionality
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabStrip_Click()
   If tabStrip.SelectedItem.Index = mintCurFrame Then Exit Sub
   ' Otherwise, hide old frame, show new.
   fraTabStripContainer(tabStrip.SelectedItem.Index).Visible = True
   fraTabStripContainer(mintCurFrame).Visible = False
   ' Set mintCurFrame to new value.
   mintCurFrame = tabStrip.SelectedItem.Index
End Sub

