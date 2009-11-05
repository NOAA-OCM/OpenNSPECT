VERSION 5.00
Begin VB.Form frmAgreeOptions 
   Caption         =   "DEM Reconditioning"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3375
      TabIndex        =   2
      Top             =   2205
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2325
      TabIndex        =   1
      Top             =   2205
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Define Reconditioning Parameters "
      Height          =   1980
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   4305
      Begin VB.TextBox txtSharpDist 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Text            =   "10"
         Top             =   1350
         Width           =   735
      End
      Begin VB.TextBox txtSmoothDist 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Text            =   "10"
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox txtBuffDist 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Text            =   "5"
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Sharp Drop/Raise:"
         Height          =   240
         Left            =   285
         TabIndex        =   5
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Smooth Drop/Raise: "
         Height          =   240
         Left            =   270
         TabIndex        =   4
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Vector Buffer (cells): "
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   450
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAgreeOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO: REMOVE


Private m_App As IApplication

Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub

Private Sub cmdCancel_Click()
    Dim intvbYesNo
    
    intvbYesNo = MsgBox("Are you sure you want to exit.  Changes will be loast?", vbYesNo, "Exit?")
    
    If intvbYesNo = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
    
End Sub

Private Sub cmdOK_Click()
    
    If Len(Trim(txtBuffDist.Text)) <> 0 And IsNumeric(CInt(txtBuffDist.Text)) Then
        g_intBuffDistance = CInt(txtBuffDist.Text)
    Else
        GoTo ErrorHandler
    End If
    
    If Len(Trim(txtSmoothDist.Text)) <> 0 And IsNumeric(CInt(txtSmoothDist.Text)) Then
        g_intSmoothDist = CInt(txtSmoothDist.Text)
    Else
        GoTo ErrorHandler
    End If
    
    If Len(Trim(txtSharpDist.Text)) <> 0 And IsNumeric(CInt(txtSharpDist.Text)) Then
        g_intSharpDist = CInt(txtSharpDist.Text)
    Else
        GoTo ErrorHandler
    End If

    g_boolAgree = True
    g_boolParams = True
    
    Call frmNewWSDelin.CheckEnabled
    
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Numeric values only please.", vbCritical, "Numeric Values Only"
    g_boolAgree = False

End Sub

