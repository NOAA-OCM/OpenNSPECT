VERSION 5.00
Begin VB.Form frmSoils 
   Caption         =   "Soils"
   ClientHeight    =   2400
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Soils Configuration  "
      Height          =   1770
      Left            =   195
      TabIndex        =   2
      Top             =   15
      Width           =   4275
      Begin VB.TextBox txtSoilsKGrid 
         Height          =   300
         Left            =   1125
         TabIndex        =   5
         Top             =   1290
         Width           =   2910
      End
      Begin VB.TextBox txtSoilsGrid 
         Height          =   300
         Left            =   1125
         TabIndex        =   4
         Top             =   795
         Width           =   2880
      End
      Begin VB.ComboBox cboSoils 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   315
         Width           =   2145
      End
      Begin VB.Label Label5 
         Caption         =   "Soils K Grid:"
         Height          =   285
         Left            =   105
         TabIndex        =   8
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Soils GRID:"
         Height          =   270
         Left            =   105
         TabIndex        =   7
         Top             =   825
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Height          =   375
      Left            =   2295
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuNew 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSoilsHelp 
         Caption         =   "Soils..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmSoils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsSoilsLoad As ADODB.Recordset
Dim rsSoilsCboClick As ADODB.Recordset
Dim rsSoilsDelete As ADODB.Recordset
Public m_App As IApplication

Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub


Private Sub cboSoils_Click()
    
    Dim strSQLSoils As String
    Set rsSoilsCboClick = New ADODB.Recordset
    
    strSQLSoils = "SELECT * FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"
    rsSoilsCboClick.CursorLocation = adUseClient
    rsSoilsCboClick.Open strSQLSoils, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'Populate the controls...
    txtSoilsGrid.Text = rsSoilsCboClick!SoilsFileName
    txtSoilsKGrid.Text = rsSoilsCboClick!SoilsKFileName
    
End Sub

Private Sub cmdQuit_Click()

    Unload frmSoils
    
End Sub

Private Sub cmdSave_Click()
    Unload frmSoils
End Sub

Private Sub Form_Load()
   'load data
   modUtil.InitComboBox cboSoils, "SOILS"
    
End Sub

Private Sub mnuDelete_Click()
    
   Dim intAns
   Dim strSQLSoilsDel As String
   Dim cntrl As Control
   
   strSQLSoilsDel = "SELECT * FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"
      
   If Not (cboSoils.Text = "") Then
   intAns = MsgBox("Are you sure you want to delete the soils setup '" & cboSoils.List(cboSoils.ListIndex) & "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   'code to handle response
        If intAns = vbYes Then
            
            'Set up a delete rs and get rid of it
            Set rsSoilsDelete = New ADODB.Recordset
            rsSoilsDelete.CursorLocation = adUseClient
            rsSoilsDelete.Open strSQLSoilsDel, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
            
            rsSoilsDelete.Delete adAffectCurrent
            rsSoilsDelete.Update
            MsgBox cboSoils.List(cboSoils.ListIndex) & " deleted.", vbOKOnly, "Record Deleted"
            
            'Clear everything, clean up form
            cboSoils.Clear
            
            txtSoilsGrid.Text = ""
            txtSoilsKGrid.Text = ""
              
            modUtil.InitComboBox cboSoils, "SOILS"
                  
            frmSoils.Refresh
                  
        ElseIf intAns = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "Please select a Soils Setup", vbCritical, "No Soils Setup Selected"
    End If

    
End Sub

Private Sub mnuNew_Click()
  
    frmSoilsSetup.init m_App
    frmSoilsSetup.Show vbModal, Me
 
End Sub



   
Private Sub mnuSoilsHelp_Click()

    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "soils.htm"

End Sub
