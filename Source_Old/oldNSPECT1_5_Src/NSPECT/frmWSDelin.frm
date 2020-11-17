VERSION 5.00
Begin VB.Form frmWSDelin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Watershed Delineations"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7560
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Browse Watershed Delineations  "
      Height          =   3450
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   7260
      Begin VB.TextBox txtLSGrid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2670
         TabIndex        =   17
         Top             =   2895
         Width           =   4395
      End
      Begin VB.ComboBox cboDEMUnits 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmWSDelin.frx":0000
         Left            =   2670
         List            =   "frmWSDelin.frx":000A
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1005
         Width           =   2550
      End
      Begin VB.ComboBox cboWSSize 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmWSDelin.frx":001C
         Left            =   2670
         List            =   "frmWSDelin.frx":0029
         TabIndex        =   15
         Text            =   "cboWSSize"
         Top             =   1710
         Width           =   1800
      End
      Begin VB.CheckBox chkHydroCorr 
         Caption         =   "Hydrologically Corrected DEM"
         Enabled         =   0   'False
         Height          =   273
         Left            =   2685
         TabIndex        =   7
         Top             =   1395
         Width           =   2587
      End
      Begin VB.TextBox txtFlowAccumGrid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2670
         TabIndex        =   6
         Top             =   2505
         Width           =   4395
      End
      Begin VB.TextBox txtWSFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2670
         TabIndex        =   5
         Top             =   2100
         Width           =   4395
      End
      Begin VB.TextBox txtStream 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2670
         TabIndex        =   4
         Top             =   1725
         Visible         =   0   'False
         Width           =   3016
      End
      Begin VB.ComboBox cboWSDelin 
         Height          =   315
         ItemData        =   "frmWSDelin.frx":0043
         Left            =   2670
         List            =   "frmWSDelin.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDEMFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2670
         TabIndex        =   2
         Top             =   645
         Width           =   4350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LS Grid:"
         Height          =   195
         Index           =   6
         Left            =   1860
         TabIndex        =   18
         Top             =   2910
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DEM Grid:"
         Height          =   195
         Index           =   12
         Left            =   1695
         TabIndex        =   14
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Units:"
         Height          =   195
         Index           =   2
         Left            =   2025
         TabIndex        =   13
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stream Agreement Layer:"
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   12
         Top             =   1725
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Watershed Delineation Name:"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   11
         Top             =   285
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Watershed:"
         Height          =   195
         Index           =   4
         Left            =   1605
         TabIndex        =   10
         Top             =   2100
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flow Accumulation Grid:"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   9
         Top             =   2505
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subwatershed Size:"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   8
         Top             =   1725
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   3645
      Width           =   975
   End
   Begin VB.Menu mnuDefWSDelin 
      Caption         =   "&Options"
      Begin VB.Menu mnuNewWSDelin 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuNewExist 
         Caption         =   "New from existing data..."
      End
      Begin VB.Menu mnuDelWSDelin 
         Caption         =   "&Delete..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWSDelin 
         Caption         =   "Watershed Delineations..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmWSDelin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmWSDelin
' *************************************************************************************
' *  Description: Form for browsing and maintaining watershed delineation
' *  Scenarios within NSPECT
' *
' *  Called By:  clsWSDelin
' *************************************************************************************

Option Explicit

Dim rsWSDelinLoad As ADODB.Recordset        'Recordset on load event
Dim rsWSDelinCboClick As ADODB.Recordset    'Recordset to pop controls
Dim rsWSDelinDelete As ADODB.Recordset      'Recordset for deletions

Private m_App As IApplication
Public m_pInputPrecipDS As IRasterDataset

' Variables used by the Error handler function

Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub


Private Sub cboWSDelin_Click()
  
    'String and recordset
    Dim strSQLPrecip As String
    Set rsWSDelinCboClick = New ADODB.Recordset
    
    'Selection based on combo box
    strSQLPrecip = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
    rsWSDelinCboClick.CursorLocation = adUseClient
    rsWSDelinCboClick.Open strSQLPrecip, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'Check for records
    If rsWSDelinCboClick.RecordCount > 0 Then
        'Populate the controls...
        txtDEMFile.Text = rsWSDelinCboClick!DEMFileName
        cboDEMUnits.ListIndex = rsWSDelinCboClick!DEMGridUnits
        txtStream.Text = rsWSDelinCboClick!StreamFileName & ""
        chkHydroCorr.Value = rsWSDelinCboClick!HydroCorrected
        cboWSSize.ListIndex = rsWSDelinCboClick!SubWSSize
        txtWSFile.Text = rsWSDelinCboClick!wsfilename & ""
        txtFlowAccumGrid.Text = rsWSDelinCboClick!FlowAccumFileName & ""
        txtLSGrid.Text = rsWSDelinCboClick!LSFileName & ""
        'Clean it
        Set rsWSDelinCboClick = Nothing

    Else
        MsgBox "Warning: There are no watershed delineation scenarios remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        Set rsWSDelinCboClick = Nothing
        Exit Sub
    End If
End Sub

Private Sub cmdQuit_Click()
   
   Unload Me

End Sub

Private Sub Form_Load()
   
   modUtil.InitComboBox cboWSDelin, "WSDELINEATION"
   
End Sub

Private Sub mnuNewExist_Click()
    frmUserWShed.init m_App
    frmUserWShed.Show vbModal, Me
End Sub

Private Sub mnuNewWSDelin_Click()

   frmNewWSDelin.init m_App
   frmNewWSDelin.Show vbModal, Me

End Sub

Private Sub mnuDelWSDelin_Click()  'Delete menu item

    Dim intAns
    Dim strSQLWSDel As String
    Dim fldWSDelin As FileSystemObject
    Dim strFolder As String
    
    strSQLWSDel = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
      
    If Not (cboWSDelin.Text = "") Then
    intAns = MsgBox("Are you sure you want to delete the watershed delineation scenario '" & cboWSDelin.List(cboWSDelin.ListIndex) & "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
    'code to handle response
        If intAns = vbYes Then
            
            'Set up a delete rs and get rid of it
            Set rsWSDelinDelete = New ADODB.Recordset
            rsWSDelinDelete.CursorLocation = adUseClient
            rsWSDelinDelete.Open strSQLWSDel, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
            
            rsWSDelinDelete.Delete adAffectCurrent
            rsWSDelinDelete.Update
            
            Set fldWSDelin = CreateObject("Scripting.FileSystemObject")
            strFolder = modUtil.g_nspectPath & "\wsdelin\" & cboWSDelin.Text
            If fldWSDelin.FolderExists(strFolder) Then
                fldWSDelin.DeleteFolder strFolder, True
            End If
            
            'Confirm
            MsgBox cboWSDelin.List(cboWSDelin.ListIndex) & " deleted.", vbOKOnly, "Record Deleted"
            
            'Clear everything, clean up form
            cboWSDelin.Clear
            chkHydroCorr.Value = 0
            txtDEMFile.Text = ""
            txtStream.Text = ""
            txtWSFile.Text = ""
            txtFlowAccumGrid.Text = ""
              
            modUtil.InitComboBox cboWSDelin, "WSDELINEATION"
                  
            frmWSDelin.Refresh
                  
        ElseIf intAns = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "Please select a watershed delineation", vbCritical, "No Scenario Selected"
    End If
    
    'cleanup
    Set fldWSDelin = Nothing
    rsWSDelinDelete.Close
    Set rsWSDelinDelete = Nothing
 

End Sub


Private Sub mnuWSDelin_Click()

    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "wsdelin.htm"

End Sub
