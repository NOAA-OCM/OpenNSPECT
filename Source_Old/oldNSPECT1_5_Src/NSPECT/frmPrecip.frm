VERSION 5.00
Begin VB.Form frmPrecip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precipitation Scenarios"
   ClientHeight    =   4170
   ClientLeft      =   2880
   ClientTop       =   3690
   ClientWidth     =   7155
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Choose a precipitation scenario to view or edit  "
      Height          =   3540
      Left            =   180
      TabIndex        =   6
      Top             =   30
      Width           =   6825
      Begin VB.ComboBox cboScenName 
         Height          =   315
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   2145
      End
      Begin VB.TextBox txtRainingDays 
         Height          =   286
         Left            =   5490
         TabIndex        =   17
         Top             =   2595
         Width           =   690
      End
      Begin VB.ComboBox cboTimePeriod 
         Height          =   315
         ItemData        =   "frmPrecip.frx":0000
         Left            =   2085
         List            =   "frmPrecip.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2581
         Width           =   2145
      End
      Begin VB.ComboBox cboPrecipType 
         Height          =   315
         ItemData        =   "frmPrecip.frx":001D
         Left            =   2085
         List            =   "frmPrecip.frx":002D
         TabIndex        =   8
         Top             =   3030
         Width           =   2145
      End
      Begin VB.CommandButton cmdBrowseFile 
         Height          =   315
         Left            =   5970
         Picture         =   "frmPrecip.frx":0055
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   2085
         TabIndex        =   0
         Top             =   805
         Width           =   3825
      End
      Begin VB.TextBox txtPrecipFile 
         Height          =   315
         Left            =   2085
         TabIndex        =   2
         Top             =   1250
         Width           =   3825
      End
      Begin VB.ComboBox cboGridUnits 
         Height          =   315
         ItemData        =   "frmPrecip.frx":08CB
         Left            =   2085
         List            =   "frmPrecip.frx":08D5
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1695
         Width           =   2130
      End
      Begin VB.ComboBox cboPrecipUnits 
         Height          =   315
         ItemData        =   "frmPrecip.frx":08E7
         Left            =   2085
         List            =   "frmPrecip.frx":08F1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2140
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scenario Name:"
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   19
         Top             =   375
         Width           =   1530
      End
      Begin VB.Label lblRainingDays 
         Caption         =   "Raining Days: "
         Height          =   315
         Left            =   4410
         TabIndex        =   16
         Top             =   2625
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   705
         TabIndex        =   14
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Precipitation Units:"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   2155
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   285
         Index           =   1
         Left            =   585
         TabIndex        =   12
         Top             =   820
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Units:"
         Height          =   285
         Index           =   6
         Left            =   810
         TabIndex        =   11
         Top             =   1710
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Precipitation Grid:"
         Height          =   285
         Index           =   7
         Left            =   465
         TabIndex        =   10
         Top             =   1265
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Period:"
         Height          =   255
         Left            =   705
         TabIndex        =   9
         Top             =   2611
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4710
      TabIndex        =   5
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5775
      TabIndex        =   7
      Top             =   3675
      Width           =   975
   End
   Begin VB.Menu mnuPrecipOpts 
      Caption         =   "Options"
      Begin VB.Menu mnuNewPrecip 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuDelPrecip 
         Caption         =   "Delete..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuPrecipHelp 
         Caption         =   "Precipitation Scenarios..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmPrecip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmPrecip
' *************************************************************************************
' *  Description: Form for browsing and maintaining precipitation scenarios
' *  within NSPECT
' *
' *  Called By:  clsPrecip
' *************************************************************************************


Option Explicit

Private rsPrecipCboClick As ADODB.Recordset
Private rsPrecipDelete As ADODB.Recordset
Private m_App As IApplication
Private boolChange As Boolean
Private boolLoad As Boolean

Public m_pInputPrecipDS As IRasterDataset


Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub


Private Sub cboGridUnits_Click()
    
    EnableSave
    
End Sub

Private Sub cboPrecipType_Click()
    If Not boolLoad Then
        EnableSave
    End If
    
End Sub

Private Sub cboPrecipUnits_Click()
    
    EnableSave
    
End Sub

Private Sub cboScenName_Click()
    
    Dim strSQLPrecip As String
    Set rsPrecipCboClick = New ADODB.Recordset
    
    strSQLPrecip = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"
    rsPrecipCboClick.CursorLocation = adUseClient
    rsPrecipCboClick.Open strSQLPrecip, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'Populate the controls...
    txtDesc.Text = rsPrecipCboClick!Description
    txtPrecipFile.Text = rsPrecipCboClick!PrecipFileName
    
    cboGridUnits.ListIndex = CInt(rsPrecipCboClick!PrecipGridUnits)
    cboPrecipUnits.ListIndex = CInt(rsPrecipCboClick!PrecipUnits)
    cboTimePeriod.ListIndex = rsPrecipCboClick!Type
    cboPrecipType.ListIndex = rsPrecipCboClick!PrecipType
    
    If rsPrecipCboClick!Type = 0 Then
        txtRainingDays.Text = rsPrecipCboClick!RainingDays
    End If
    
    cmdSave.Enabled = False
    
     
End Sub

Private Sub cboTimePeriod_Click()
    
    If cboTimePeriod.ListIndex = 0 Then
        lblRainingDays.Visible = True
        txtRainingDays.Visible = True
    Else
        lblRainingDays.Visible = False
        txtRainingDays.Visible = False
    End If
    
End Sub

Private Sub cmdBrowseFile_Click()
    
    Dim pPrecipRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    
On Error GoTo ErrHandler:

    Set m_pInputPrecipDS = AddInputFromGxBrowserText(txtPrecipFile, "Choose Precipitation GRID", frmPrecip, 0)
    
    If m_pInputPrecipDS Is Nothing Then
        Exit Sub
    Else
    
    Set pPrecipRasterDataset = m_pInputPrecipDS
    
        If CheckSpatialReference(pPrecipRasterDataset) Is Nothing Then
            
            MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
            txtPrecipFile.Text = ""
            Exit Sub
        
        Else
        
            Set pProjCoord = CheckSpatialReference(pPrecipRasterDataset)
            Set pDistUnit = pProjCoord.CoordinateUnit
            intUnit = pDistUnit.MetersPerUnit
            
                If intUnit = 1 Then
                    cboPrecipUnits.ListIndex = 0
                Else
                    cboPrecipUnits.ListIndex = 1
                End If
            
            cboPrecipUnits.Refresh
        End If
    
    End If
    
    
    Set pPrecipRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing

Exit Sub
ErrHandler:
    Exit Sub
    
End Sub
    
Private Sub cmdSave_Click()

    If CheckParams = True Then
    
        rsPrecipCboClick!Name = cboScenName.Text
        rsPrecipCboClick!Description = txtDesc.Text
        rsPrecipCboClick!PrecipFileName = txtPrecipFile.Text
        rsPrecipCboClick!PrecipGridUnits = cboGridUnits.ListIndex
        rsPrecipCboClick!PrecipUnits = cboPrecipUnits.ListIndex
        rsPrecipCboClick!PrecipType = cboPrecipType.ListIndex
        rsPrecipCboClick!Type = cboTimePeriod.ListIndex
        
        If cboTimePeriod.ListIndex = 0 Then
            rsPrecipCboClick!RainingDays = CInt(txtRainingDays.Text)
        Else
            rsPrecipCboClick!RainingDays = 0
        End If
        
        rsPrecipCboClick.Update
        boolChange = False
        
        MsgBox cboScenName.Text & " saved successfully.", vbOKOnly, "Record Saved"
        Unload frmPrecip
        
        Set rsPrecipCboClick = Nothing

    Else
    
        Exit Sub
    
    End If
    
End Sub



Private Sub Form_Load()
   
   'load data
   boolLoad = True
   modUtil.InitComboBox cboScenName, "PRECIPSCENARIO"
   cmdSave.Enabled = False
   boolChange = False
   boolLoad = False
   
End Sub

Private Sub cmdQuit_Click()
   
    Dim intSave
   
    If boolChange Then
        intSave = MsgBox("You have made changes to this record, are you sure you want to quit?", vbYesNo, "Quit?")
    
            If intSave = vbYes Then
                Unload frmPrecip
            ElseIf intSave = vbNo Then
                
                Exit Sub
            End If
    Else
        Unload Me
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    rsPrecipCboClick.Close
    Set rsPrecipCboClick = Nothing

End Sub

Private Sub mnuNewPrecip_Click()

   frmNewPrecip.Show vbModal, Me

End Sub

Private Sub mnuDelPrecip_Click()

   Dim intAns
   Dim strSQLPrecipDel As String
   Dim cntrl As Control
   
   strSQLPrecipDel = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"
      
   If Not (cboScenName.Text = "") Then
   intAns = MsgBox("Are you sure you want to delete the precipitation scenario '" & cboScenName.List(cboScenName.ListIndex) & "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   'code to handle response
        If intAns = vbYes Then
            
            'Set up a delete rs and get rid of it
            Set rsPrecipDelete = New ADODB.Recordset
            rsPrecipDelete.CursorLocation = adUseClient
            rsPrecipDelete.Open strSQLPrecipDel, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
            
            rsPrecipDelete.Delete adAffectCurrent
            rsPrecipDelete.Update
            MsgBox cboScenName.List(cboScenName.ListIndex) & " deleted.", vbOKOnly, "Record Deleted"
            
            'Clear everything, clean up form
            cboScenName.Clear
            'cboGridUnits.Clear
            'cboPrecipUnits.Clear
            
            txtDesc.Text = ""
            'txtDuration.Text = ""
            txtPrecipFile.Text = ""
              
            modUtil.InitComboBox cboScenName, "PRECIPSCENARIO"
                  
            frmPrecip.Refresh
                  
        ElseIf intAns = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "Please select a Precipitation Scenario", vbCritical, "No Scenario Selected"
    End If
 

End Sub
        
Private Sub mnuPrecipHelp_Click()

    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "precip.htm"

End Sub

Private Sub optAnnual_Click()
    EnableSave
End Sub

Private Sub txtDesc_Change()
    
    EnableSave
    txtDesc.Text = Replace(txtDesc.Text, "'", "")

End Sub

Private Sub EnableSave()
    
    cmdSave.Enabled = True
    boolChange = True
    
End Sub

Private Sub txtDuration_Change()

    EnableSave
   
End Sub

Private Sub txtPrecipFile_Change()

    EnableSave
    
End Sub

Private Function CheckParams() As Boolean

    'Check the inputs of the form, before saving
    If Len(txtDesc.Text) = 0 Then
        MsgBox "Please enter a description for this scenario", vbCritical, "Description Missing"
        txtDesc.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If txtPrecipFile.Text = " " Or txtPrecipFile.Text = "" Then
        MsgBox "Please select a valid precipitation GRID before saving.", vbCritical, "GRID Missing"
        txtPrecipFile.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboGridUnits.Text = "" Then
        MsgBox "Please select GRID units.", vbCritical, "Units Missing"
        cboGridUnits.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboPrecipUnits.Text = "" Then
        MsgBox "Please select precipitation units.", vbCritical, "Units Missing"
        cboPrecipUnits.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If Len(cboPrecipType.Text) = 0 Then
        MsgBox "Please select a Precipitation Type.", vbCritical, "Precipitation Type Missing"
        cboPrecipType.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If Len(cboTimePeriod.Text) = 0 Then
        MsgBox "Please select a Time Period.", vbCritical, "Precipitation Time Period Missing"
        cboTimePeriod.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboTimePeriod.ListIndex = 0 Then
        If Not IsNumeric(txtRainingDays.Text) Or Len(txtRainingDays.Text) = 0 Then
            MsgBox "Please enter a numeric value for Raining Days.", vbCritical, "Raining Days Value Incorrect"
            txtRainingDays.SetFocus
            CheckParams = False
            Exit Function
        End If
    End If
    
    'if it got through all that, then set it to true
    CheckParams = True
    

End Function

Private Sub txtRainingDays_Change()
    
    EnableSave
    
End Sub
