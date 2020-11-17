VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPollutants 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pollutants"
   ClientHeight    =   8490
   ClientLeft      =   1695
   ClientTop       =   1905
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   1005
      Top             =   7935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtActiveCellWQStd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3135
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   7965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtActiveCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1800
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   7965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7245
      Left            =   195
      TabIndex        =   4
      Top             =   570
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   12779
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   7056
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Coefficients"
      TabPicture(0)   =   "frmPollutants.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grdPollDef"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboCoeffSet"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCoeffSetDesc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtLCType"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Water Quality Standards"
      TabPicture(1)   =   "frmPollutants.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdWQStd"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtLCType 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5475
         TabIndex        =   7
         Top             =   480
         Width           =   1665
      End
      Begin VB.TextBox txtCoeffSetDesc 
         Height          =   285
         Left            =   1500
         TabIndex        =   6
         Top             =   900
         Width           =   5625
      End
      Begin VB.ComboBox cboCoeffSet 
         Height          =   315
         ItemData        =   "frmPollutants.frx":0038
         Left            =   1500
         List            =   "frmPollutants.frx":003A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWQStd 
         Height          =   4320
         Left            =   -74640
         TabIndex        =   16
         Top             =   720
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   7620
         _Version        =   393216
         Cols            =   5
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPollDef 
         Height          =   5535
         Left            =   195
         TabIndex        =   17
         Top             =   1560
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   8
         BackColorSel    =   12648447
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Label Label2 
         Caption         =   "Threshold Units: ug/L"
         Height          =   255
         Left            =   -74520
         TabIndex        =   18
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   1290
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coefficients (mg/L)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   12
         Top             =   1290
         Width           =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   11
         Top             =   1290
         Width           =   3450
      End
      Begin VB.Label Label1 
         Caption         =   "Land Cover Type:"
         Height          =   255
         Index           =   7
         Left            =   4065
         TabIndex        =   10
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Index           =   6
         Left            =   270
         TabIndex        =   9
         Top             =   885
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Coefficient Set:"
         Height          =   255
         Index           =   5
         Left            =   255
         TabIndex        =   8
         Top             =   510
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5565
      TabIndex        =   2
      Top             =   7935
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6630
      TabIndex        =   3
      Top             =   7935
      Width           =   975
   End
   Begin VB.ComboBox cboPollName 
      Height          =   315
      ItemData        =   "frmPollutants.frx":003C
      Left            =   1590
      List            =   "frmPollutants.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   2250
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   300
      Top             =   7920
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Pollutant Name:"
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   165
      Width           =   1335
   End
   Begin VB.Menu mnuPoll 
      Caption         =   "&Pollutants"
      Index           =   1
      Begin VB.Menu mnuAddPoll 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuDeletePoll 
         Caption         =   "&Delete..."
      End
   End
   Begin VB.Menu mnuCoeff 
      Caption         =   "&Coefficients"
      Begin VB.Menu mnuCoeffNewSet 
         Caption         =   "&New Set..."
      End
      Begin VB.Menu mnuCoeffCopySet 
         Caption         =   "&Copy Set..."
      End
      Begin VB.Menu mnuCoeffDeleteSet 
         Caption         =   "&Delete Set..."
      End
      Begin VB.Menu mnuCoeffImportSet 
         Caption         =   "&Import Set..."
      End
      Begin VB.Menu mnuCoeffExportSet 
         Caption         =   "&Export Set..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuPollHelp 
         Caption         =   "Pollutants..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuCoeffHelp 
         Caption         =   "Coefficients..."
         Shortcut        =   +{F2}
      End
   End
End
Attribute VB_Name = "frmPollutants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmPollutants
' *************************************************************************************
' *  Description: Form for browsing pollutants
' *  within NSPECT
' *
' *  Called By:  NSPECT clsPollutants
' *************************************************************************************
' *  Subs/Functions:
' *     ExportCoeffSet(strFileName As String): Exports data to textfile
' *     KeyMoveUpdate:  Handles <- -> ^ key input from user while in data grid
' *     ValidateGridValues:  Function returns Boolean value whether or not all data is
' *         ok before updating..specifically if coeffs are numeric and in range 0 - 100
' *
' *  Misc:
' *************************************************************************************

Option Explicit

Private m_App As IApplication

Private rsPollCboClick As ADODB.Recordset   'RS on cboPollName click event
Private rsCoeff As ADODB.Recordset
Private rsLCType As ADODB.Recordset
Private rsFullCoeff As ADODB.Recordset      'RS on cboCoeffSet click event
Private rsCoeffs As ADODB.Recordset         'RS that fills FlexGrid
Private rsWQStds As ADODB.Recordset         'Rs that fills FelexGrid on Standards Tab

Private boolLoaded As Boolean
Private boolChanged As Boolean              'Bool for enabling save button
Private boolDescChanged As Boolean          'Boolship for seeing if Description Changed
Private boolSaved As Boolean                'Boolship for whether or not things have saved

Private m_intCurFrame As Integer            'Current Frame visible in SSTab
Private m_intPollRow As Integer             'Row Number for grdPolldef
Private m_intPollCol As Integer             'Column Number for grdPollDef
Private intRowWQ As Integer                 'Row Number for grdPollWQStd
Private intColWQ As Integer                 'Column Number for grdPollDef
Private m_intPollID As Integer              'There's a need to have the PollID so we'll store it here
Private m_intLCTypeID As Integer            'Land Class (CCAP) ID - needed to add new coefficient sets
Private m_intCoeffID As Integer             'Key for CoefficientSetID - needed to add new coefficients 'See above

Private m_strUndoText As String             'Text for txtActiveCell     |
Private m_strUndoTextWQ As String           'Text for txtActiveCellWQ   |-all three used to detect change
Private m_strUndoDesc As String             'Text for Description       |
Private m_strLCType As String               'Need for name, we'll store here


' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmPollutants.frm"


Private Sub cmdSave_Click()
  On Error GoTo ErrorHandler


    If ValidateGridValues Then
        
        UpdateValues
        boolSaved = True
        MsgBox cboPollName.Text & " saved successfully.", vbInformation, "N-SPECT"
        Unload Me
        
    End If

    

  Exit Sub
ErrorHandler:
  HandleError True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

   
   boolChanged = False
   
   'Initialize the two flex Grids
   InitPollDefGrid grdPollDef
   InitPollWQStdGrid grdWQStd
   
   'Toss in the names of all pollutants and call the cbo click event
   InitComboBox cboPollName, "Pollutant"
   
   SSTab1.Tab = 0
   boolLoaded = True
   boolChanged = False
   boolSaved = False
   

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboPollName_Click()
    
    Dim strSQLPollutant As String
    Dim strSQLLCType As String
    Dim strSQLCoeff As String
    Dim strSQLWQStd As String
    
On Error GoTo ErrorHandler


'Check to see if things have changed
If boolChanged Then

    intYesNo = MsgBox(strYesNo, vbYesNo, strYesNoTitle)
    
    If intYesNo = vbYes Then
    
        UpdateValues
        
        boolChanged = False
              
        Set rsPollCboClick = New ADODB.Recordset
         
        'Selection based on combo box
        strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
        rsPollCboClick.Open strSQLPollutant, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        m_intPollID = rsPollCboClick!POLLID
         
        Set rsCoeff = New ADODB.Recordset
         
        strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick!POLLID & ""
        rsCoeff.Open strSQLCoeff, modUtil.g_ADOConn, adOpenKeyset, adLockOptimistic
    
        Set rsLCType = New ADODB.Recordset
         
        strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff!LCTypeID & ""
        rsLCType.Open strSQLLCType, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        m_strLCType = rsLCType!Name
        m_intLCTypeID = rsLCType!LCTypeID
         
        'Fill everything based on that
        cboCoeffSet.Clear
        Do Until rsCoeff.EOF
           cboCoeffSet.AddItem rsCoeff!Name
           rsCoeff.MoveNext
        Loop
             
        cboCoeffSet.ListIndex = 0
         
        txtLCType.Text = m_strLCType
        
        'Fill the Water Quality Standards Tab
        Set rsWQStds = New ADODB.Recordset
         
        strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & _
        "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & _
        "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & _
        "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick!POLLID
            
        rsWQStds.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        Set grdWQStd.Recordset = rsWQStds
         
        'Cleanup
        Set rsPollCboClick = Nothing
        'Set rsCoeff = Nothing
        Set rsLCType = Nothing
        Set rsWQStds = Nothing
    
    ElseIf intYesNo = vbNo Then
     
        boolChanged = False
              
        Set rsPollCboClick = New ADODB.Recordset
         
        'Selection based on combo box
        strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
        rsPollCboClick.Open strSQLPollutant, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        m_intPollID = rsPollCboClick!POLLID
         
        Set rsCoeff = New ADODB.Recordset
         
        strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick!POLLID & ""
        rsCoeff.Open strSQLCoeff, modUtil.g_ADOConn, adOpenKeyset, adLockOptimistic
    
        Set rsLCType = New ADODB.Recordset
         
        strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff!LCTypeID & ""
        rsLCType.Open strSQLLCType, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        m_strLCType = rsLCType!Name
        m_intLCTypeID = rsLCType!LCTypeID
         
        'Fill everything based on that
        cboCoeffSet.Clear
        Do Until rsCoeff.EOF
           cboCoeffSet.AddItem rsCoeff!Name
           rsCoeff.MoveNext
        Loop
             
        cboCoeffSet.ListIndex = 0
         
        txtLCType.Text = m_strLCType
        
        'Fill the Water Quality Standards Tab
        Set rsWQStds = New ADODB.Recordset
         
        strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & _
        "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & _
        "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & _
        "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick!POLLID
            
        rsWQStds.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        Set grdWQStd.Recordset = rsWQStds
         
        'Cleanup
        Set rsPollCboClick = Nothing
        'Set rsCoeff = Nothing
        Set rsLCType = Nothing
        Set rsWQStds = Nothing
        
    End If
    
Else
         
     boolChanged = False
              
     Set rsPollCboClick = New ADODB.Recordset
     
     'Selection based on combo box
     strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
     rsPollCboClick.Open strSQLPollutant, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
     m_intPollID = rsPollCboClick!POLLID
     
     Set rsCoeff = New ADODB.Recordset
     
     strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick!POLLID & ""
     rsCoeff.Open strSQLCoeff, modUtil.g_ADOConn, adOpenKeyset, adLockOptimistic

     Set rsLCType = New ADODB.Recordset
     
     strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff!LCTypeID & ""
     rsLCType.Open strSQLLCType, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
     m_strLCType = rsLCType!Name
     m_intLCTypeID = rsLCType!LCTypeID
     
     'Fill everything based on that
     cboCoeffSet.Clear
     Do Until rsCoeff.EOF
        cboCoeffSet.AddItem rsCoeff!Name
        rsCoeff.MoveNext
     Loop
         
     cboCoeffSet.ListIndex = 0
     
     txtLCType.Text = m_strLCType
    
     'Fill the Water Quality Standards Tab
     Set rsWQStds = New ADODB.Recordset
     
     strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & _
     "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & _
     "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & _
     "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick!POLLID
        
     rsWQStds.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
     Set grdWQStd.Recordset = rsWQStds
     
     'Cleanup
     Set rsPollCboClick = Nothing
     'Set rsCoeff = Nothing
     Set rsLCType = Nothing
     Set rsWQStds = Nothing
    
End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "cboPollName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboCoeffSet_Click()
  On Error GoTo ErrorHandler

    
txtActiveCell.Visible = False
    
If boolChanged Then

    intYesNo = MsgBox(strYesNo, vbYesNo, strYesNoTitle)
    
    If intYesNo = vbYes Then
    
        If ValidateGridValues Then
            UpdateValues
            boolChanged = False
        End If
     Else
            GoTo notSave
     End If

Else
        
notSave:
    boolChanged = False
    grdPollDef.Clear
    
    Dim strSQLFullCoeff As String
    Set rsFullCoeff = New ADODB.Recordset
    
    strSQLFullCoeff = "SELECT COEFFICIENTSET.NAME, COEFFICIENTSET.DESCRIPTION, " & _
                      "COEFFICIENTSET.COEFFSETID, LCTYPE.NAME as NAME2 " & _
                      "FROM COEFFICIENTSET INNER JOIN LCTYPE " & _
                      "ON COEFFICIENTSET.LCTYPEID = LCTYPE.LCTYPEID " & _
                      "WHERE COEFFICIENTSET.NAME LIKE '" & cboCoeffSet.Text & "'"

    rsFullCoeff.Open strSQLFullCoeff, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    Debug.Print rsFullCoeff.RecordCount

    With txtCoeffSetDesc
        .Text = rsFullCoeff!Description & ""
        .Refresh
    End With
    
    txtLCType.Text = rsFullCoeff!Name2 & ""
    
        
    Dim strSQLCoeffs As String
    
    strSQLCoeffs = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.Coeff1 As Type1, COEFFICIENT.Coeff2 as Type2, " & _
                   "COEFFICIENT.Coeff3 as Type3, COEFFICIENT.Coeff4 as Type4, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & _
                   "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & _
                   "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & _
                   "WHERE COEFFICIENT.COEFFSETID = " & rsFullCoeff!CoeffSetID & " ORDER BY LCCLASS.VALUE"
    
    Set rsCoeffs = New ADODB.Recordset
    rsCoeffs.Open strSQLCoeffs, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    Set grdPollDef.Recordset = rsCoeffs
    
    'Cleanup
    rsCoeffs.Close
    rsFullCoeff.Close
    
    Set rsFullCoeff = Nothing
    Set rsCoeffs = Nothing
    
End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "cboCoeffSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

   
    If (boolChanged Or boolDescChanged) And Not boolSaved Then
    
    intYesNo = MsgBox(strYesNo, vbYesNo, strYesNoTitle)
                
            If intYesNo = vbYes Then
            
                If ValidateGridValues Then
                    UpdateValues
                    MsgBox "Data saved successfully.", vbOKOnly, "Save Successful"
                    boolChanged = False
                    boolDescChanged = False
                    boolSaved = True
                    Unload Me
                End If
            
            Else
                
                Unload Me
            
            End If
    Else
        
        Unload Me
    End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
        

Private Sub grdPollDef_Click()
  On Error GoTo ErrorHandler

    'Code for Poll Def Grid - moves txtActiveCell to appropriate cell
    
    m_intPollRow = grdPollDef.MouseRow
    m_intPollCol = grdPollDef.MouseCol
    
    txtActiveCell.Visible = False

    'Column greater than or equal to 1 and same for row
    If m_intPollCol >= 3 Then

        With txtActiveCell
            'Put in position
            .Move SSTab1.Left + grdPollDef.Left + grdPollDef.CellLeft, _
            SSTab1.Top + grdPollDef.Top + grdPollDef.CellTop, _
            grdPollDef.CellWidth - 30, _
            grdPollDef.CellHeight - 150
             
            If m_intPollRow <> 0 Then
                .Visible = True
                m_strUndoText = grdPollDef.TextMatrix(m_intPollRow, m_intPollCol)
                .Text = grdPollDef.TextMatrix(m_intPollRow, m_intPollCol)
                .SetFocus

'                If IsNumeric(m_strUndoText) Then
'                    txtActiveCell.Alignment = 1
'                Else
'                    txtActiveCell.Alignment = 0
'                End If

             End If

        End With

    End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "grdPollDef_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub grdWQStd_Click()
  On Error GoTo ErrorHandler

    
    intRowWQ = grdWQStd.MouseRow
    intColWQ = grdWQStd.MouseCol
    
    txtActiveCellWQStd.Visible = False
    'We want to limit editing to only the threshold columns
    If intColWQ = 3 And intRowWQ >= 1 Then

        With txtActiveCellWQStd
            
            .Move SSTab1.Left + grdWQStd.Left + grdWQStd.CellLeft, _
            SSTab1.Top + grdWQStd.Top + grdWQStd.CellTop, _
            grdWQStd.CellWidth - 30, _
            grdWQStd.CellHeight - 150

            If intRowWQ <> 0 Then
                .Visible = True
                m_strUndoTextWQ = grdWQStd.TextMatrix(intRowWQ, intColWQ)
                .Text = grdWQStd.TextMatrix(intRowWQ, intColWQ)
                .SetFocus

                If IsNumeric(m_strUndoTextWQ) Then
                        txtActiveCellWQStd.Alignment = 1
                    Else
                        txtActiveCellWQStd.Alignment = 0
                End If

             End If

        End With

    End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "grdWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub mnuCoeffHelp_Click()
    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "pol_coeftab.htm"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Coefficients menu option processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCoeffNewSet_Click()
  On Error GoTo ErrorHandler

   
   g_boolAddCoeff = True
   frmAddCoeffSet.Show vbModal, Me



  Exit Sub
ErrorHandler:
  HandleError True, "mnuCoeffNewSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuCoeffDeleteSet_Click()
   
'Using straight command text to rid ourselves of the dreaded coefficient sets

On Error GoTo ErrHandler:
   Dim intAns As Integer
   intAns = MsgBox("Are you sure you want to delete the coefficient set '" & cboCoeffSet.List(cboCoeffSet.ListIndex) & "' associated with pollutant '" & cboPollName.List(cboPollName.ListIndex) & "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   
   'code to handle response
   If intAns = vbYes Then
    
        Dim strDeleteCoeffSet As String
        strDeleteCoeffSet = "DELETE * from COEFFICIENTSET WHERE NAME LIKE '" & cboCoeffSet.Text & "'"
        
        modUtil.g_ADOConn.Execute strDeleteCoeffSet
                         
        MsgBox cboCoeffSet.Text & " deleted.", vbOKOnly, "Record Deleted"
                
        cboPollName.Clear
        cboCoeffSet.Clear
        
        modUtil.InitComboBox cboPollName, "Pollutant"
        
        frmPollutants.Refresh
    
    Else
    
        Exit Sub
    
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error deleting coefficient set.", vbCritical, "Error"
    MsgBox Err.Number & ": " & Err.Description
   
End Sub

Private Sub mnuCoeffCopySet_Click()
  On Error GoTo ErrorHandler
  
    g_boolCopyCoeff = True
    frmCopyCoeffSet.init rsCoeff
    frmCopyCoeffSet.Show vbModal, Me

  Exit Sub
ErrorHandler:
  HandleError True, "mnuCoeffCopySet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuCoeffImportSet_Click()
  On Error GoTo ErrorHandler

    
    Load frmImportCoeffSet
    frmImportCoeffSet.Show vbModal, Me
    


  Exit Sub
ErrorHandler:
  HandleError True, "mnuCoeffImportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'Export Menu
Private Sub mnuCoeffExportSet_Click()
  On Error GoTo ErrorHandler

   Dim intAns As Integer
   
   'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = Replace(MSG1, "<name>", "Coefficient Set")
     .DialogTitle = Replace(MSG3, "<name>", "Coefficient Set")
     .FilterIndex = 1
     .DefaultExt = ".txt"
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
     .ShowSave
   End With
   If Len(dlgCMD1.FileName) > 0 Then
      'Export Water Quality Standard to file - dlgCMD1.FileName
      ExportCoeffSet dlgCMD1.FileName
   End If



  Exit Sub
ErrorHandler:
  HandleError True, "mnuCoeffExportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Exports your current standard and pollutants to text or csv.
Private Sub ExportCoeffSet(strFileName As String)
  On Error GoTo ErrorHandler


    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
     
    Dim rsNew As Recordset
    Dim theLine
    
    Set fl = fso.CreateTextFile(strFileName, True)

    Dim i As Integer
    
    'Write name of pollutant and threshold
    For i = 1 To grdPollDef.Rows - 1
        fl.WriteLine grdPollDef.TextMatrix(i, 1) & "," & grdPollDef.TextMatrix(i, 3) & "," & _
        grdPollDef.TextMatrix(i, 4) & "," & grdPollDef.TextMatrix(i, 5) & "," & _
        grdPollDef.TextMatrix(i, 6)
    
    Next i

    fl.Close
        
    Set fso = Nothing
    Set fl = Nothing



  Exit Sub
ErrorHandler:
  HandleError False, "ExportCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Pollutants menu option processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddPoll_Click()
  On Error GoTo ErrorHandler

    
    
    frmNewPollutants.Show vbModal, Me



  Exit Sub
ErrorHandler:
  HandleError True, "mnuAddPoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuDeletePoll_Click()
  On Error GoTo ErrorHandler


   Dim intAns As Integer
   intAns = MsgBox("Are you sure you want to delete the pollutant '" + cboPollName.List(cboPollName.ListIndex) + "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   'code to handle response
   
    If intAns = vbYes Then
        DeletePollutant cboPollName.List(cboPollName.ListIndex)
    Else
        Exit Sub
    End If



  Exit Sub
ErrorHandler:
  HandleError True, "mnuDeletePoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub mnuPollHelp_Click()
    
    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "pollutants.htm"
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  On Error GoTo ErrorHandler

'Clear the text box if user jumps from one tab to the next
    Select Case SSTab1.Tab
        Case 0
            txtActiveCellWQStd.Visible = False
        Case 1
            txtActiveCell.Visible = False
    End Select



  Exit Sub
ErrorHandler:
  HandleError True, "SSTab1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_Change()
  On Error GoTo ErrorHandler

        
        
    grdPollDef.Text = txtActiveCell.Text
    


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler


With grdPollDef
        Select Case KeyCode
            Case vbKeyEscape 'if the user pressed escape, then get out without changing
                .Text = m_strUndoText
                txtActiveCell.Visible = False
                .SetFocus
                
            Case 13 'if the user presses enter, get out of the textbox
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
            
            Case vbKeyUp 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
                .Text = txtActiveCell.Text
                If .row > 1 Then
                    .row = .row - 1
                    KeyMoveUpdate
                Else
                    If .col < 6 Then
                        .col = .col + 1
                        .row = .Rows - 1
                    Else
                        .row = .row
                    End If
                    KeyMoveUpdate
                End If
            
            Case vbKeyDown 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                If .row < grdPollDef.Rows - 1 Then
                    .row = .row + 1
                    KeyMoveUpdate
                Else
                    If .col > 1 Then
                        .row = 1 'again, if the row is on the last row, don't move cells
                        .col = .col - 1
                    Else
                        .Text = txtActiveCell.Text
                        txtActiveCell.Visible = False
                    End If
                    KeyMoveUpdate
                End If
            Case vbKeyLeft 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                txtActiveCell.Visible = False
                If .col > 3 Then
                    .col = .col - 1
                    KeyMoveUpdate
                    
                Else
                
                If .row > 1 And .col = 3 Then
                    .col = 6
                    .row = .row - 1
                    KeyMoveUpdate
                End If
                End If
            Case vbKeyRight 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                If .col < 6 Then
                    .col = .col + 1
                    KeyMoveUpdate
                Else
                    If .row < grdPollDef.Rows - 1 Then
                        .col = 3
                        .row = .row + 1
                        KeyMoveUpdate
                    End If
                End If
        End Select
    End With



  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub KeyMoveUpdate()
  On Error GoTo ErrorHandler

'This guy basically replicates the functionality of the grdpolldef_Click event and is used
'in a couple of instances for moving around the grid.
 
    m_intPollRow = grdPollDef.row
    m_intPollCol = grdPollDef.col

    txtActiveCell.Visible = False

    If (m_intPollCol >= 3 And m_intPollCol < 7) And (m_intPollRow >= 1) Then
        
        With txtActiveCell
        
             .Move SSTab1.Left + grdPollDef.Left + grdPollDef.CellLeft, _
            SSTab1.Top + grdPollDef.Top + grdPollDef.CellTop, _
            grdPollDef.CellWidth - 30, _
            grdPollDef.CellHeight - 150
             
            If m_intPollRow <> 0 Then
                
                .Visible = True
                m_strUndoText = grdPollDef.TextMatrix(m_intPollRow, m_intPollCol)
                .Text = m_strUndoText
            End If
        End With
        
        txtActiveCell.SetFocus
    
    ElseIf m_intPollCol = 0 Then
        
        txtActiveCell.Visible = False
    
    End If
    
  Exit Sub
ErrorHandler:
  HandleError False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler


If Not boolChanged Then
    
    If txtActiveCell.Text <> m_strUndoText Then
        boolChanged = True
        CmdSaveEnabled
    End If

End If


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_GotFocus()
  On Error GoTo ErrorHandler


    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)



  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub txtActiveCellWQStd_Change()
  On Error GoTo ErrorHandler

    
    grdWQStd.Text = txtActiveCellWQStd.Text
    


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCellWQStd_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Just making the entire cell text selected
Private Sub txtActiveCellWQStd_GotFocus()
  On Error GoTo ErrorHandler
    
    txtActiveCellWQStd.SelStart = 0
    txtActiveCellWQStd.SelLength = Len(txtActiveCellWQStd.Text)
    m_strUndoTextWQ = txtActiveCellWQStd.Text


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCellWQStd_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub txtActiveCellWQStd_Validate(Cancel As Boolean)

On Error GoTo ErrorHandler


If Not boolChanged Then
    
    If txtActiveCellWQStd.Text <> m_strUndoTextWQ Then
        boolChanged = True
        CmdSaveEnabled
    End If

End If


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Function ValidateGridValues() As Boolean
  On Error GoTo ErrorHandler

'Need to validate each grid value before saving.  Essentially we take it a row at a time,
'then rifle through each column of each row.  Case Select tests each each x,y value depending
'on column... 3-6 must be 1-100 range

'Returns: True or False

    Dim varActive As Variant            'txtActiveCell value
    Dim i As Integer
    Dim j As Integer
    Dim iQstd As Integer
    Dim jQstd As Integer
        
    For i = 1 To grdPollDef.Rows - 1
        
        For j = 3 To 6
    
        varActive = grdPollDef.TextMatrix(i, j)
        
        If InStr(1, CStr(varActive), ".", vbTextCompare) > 0 Then
            If (Len(Split(CStr(varActive), ".")(1)) > 4) Then
                ErrorGenerator Err6, i, j
                grdPollDef.col = j
                grdPollDef.row = i
                ValidateGridValues = False
                KeyMoveUpdate
                Exit Function
            End If
        End If
        
        If Not IsNumeric(varActive) Or (varActive < 0) Or (varActive > 1000) Then
            ErrorGenerator Err6, i, j
            grdPollDef.col = j
            grdPollDef.row = i
            ValidateGridValues = False
            KeyMoveUpdate
            Exit Function
        End If
        
        
          
        Next j
    
    Next i
        
    For iQstd = 1 To grdWQStd.Rows - 1
        varActive = grdWQStd.TextMatrix(iQstd, 3)
        
        If Not IsNumeric(varActive) Or (varActive < 0) Then
            ErrorGenerator Err5, iQstd, 3
            grdWQStd.col = 3
            grdWQStd.row = iQstd
            ValidateGridValues = False
            Exit Function
        End If
    Next iQstd
    
    ValidateGridValues = True
        
  Exit Function
ErrorHandler:
  HandleError False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Sub txtCoeffSetDesc_GotFocus()
  On Error GoTo ErrorHandler

    m_strUndoDesc = txtCoeffSetDesc.Text

  Exit Sub
ErrorHandler:
  HandleError True, "txtCoeffSetDesc_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtCoeffSetDesc_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler


'If Not boolDescChanged Then
    If m_strUndoDesc <> txtCoeffSetDesc.Text Then
        boolDescChanged = True
        CmdSaveEnabled
    End If
'End If

  Exit Sub
ErrorHandler:
  HandleError True, "txtCoeffSetDesc_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub DeletePollutant(strName As String)
  
'Can ya guess what this one does?

  On Error GoTo ErrorHandler

    Dim strPollDelete As String
    Dim strPoll2Delete As String
   
    Dim rsPollDelete As ADODB.Recordset
    'Dim rsLCClassDelete As ADODB.Recordset
   
    strPollDelete = "Delete * FROM Pollutant WHERE NAME LIKE '" & strName & "'"
      
              
    modUtil.g_ADOConn.Execute strPollDelete
                     
    MsgBox strName & " deleted.", vbOKOnly, "Record Deleted"
            
    frmPollutants.cboPollName.Clear
    modUtil.InitComboBox frmPollutants.cboPollName, "Pollutant"
    frmPollutants.Refresh

  Exit Sub
ErrorHandler:
  HandleError False, "DeletePollutant " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub AddCoefficient(strCoeffName As String, strLCType As String)
  On Error GoTo ErrorHandler

'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever
    
    Dim strNewLcType As String              'CmdString for inserting new coefficientset
    Dim strGetLcType As String
    Dim strDefault As String                '
    Dim strNewCoeff As String
    Dim strNewCoeffID As String             'Holder for the CoefficientSetID
    Dim strInsertNewCoeff As String         'Putting the newly created coefficients in their table
    Dim intCoeffSetID As Integer
    Dim i As Integer
    
    Dim rsLCType As New ADODB.Recordset
    Dim rsCopySet As ADODB.Recordset        'First RS
    Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    Dim rsNewCoeff As New ADODB.Recordset
    
    strGetLcType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCType & "'"
    rsLCType.Open strGetLcType, g_ADOConn, adOpenStatic, adLockReadOnly
    
    strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & _
        Replace(strCoeffName, "'", "''") & "'," & _
        Replace(m_intPollID, "'", "''") & "," & _
        Replace(rsLCType!LCTypeID, "'", "''") & ")"
    
    'First need to add the coefficient set to that table
    g_ADOConn.Execute strNewLcType, adAffectCurrent
    
    'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & _
        "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
        
    rsCoeffSetID.Open strNewCoeffID, g_ADOConn, adOpenKeyset, adLockReadOnly
    intCoeffSetID = rsCoeffSetID!CoeffSetID
        
        
    strDefault = "SELECT LCTYPE.LCTYPEID, LCCLASS.LCCLASSID, LCCLASS.NAME As valName, " & _
        "LCCLASS.VAlue as valValue FROM LCTYPE " & _
        "INNER JOIN LCCLASS ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & _
        "WHERE LCTYPE.Name Like " & "'" & strLCType & "' ORDER BY LCCLASS.Value"
           
       Debug.Print strDefault
    Set rsCopySet = New ADODB.Recordset
    rsCopySet.Open strDefault, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    grdPollDef.Clear
    grdPollDef.Rows = rsCopySet.RecordCount + 1
    
    
    modUtil.InitPollDefGrid grdPollDef  'Call this again to set it up as we cleared it
    
    Dim cmdInsert As New ADODB.Command
    
    'Now loopy loo to populate values.
    Dim strNewCoeff1 As String
    strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    
    rsNewCoeff.Open strNewCoeff1, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    i = 0
    
    rsCopySet.MoveFirst
    'For i = 1 To rsCopySet.RecordCount
    Do While Not rsCopySet.EOF
        i = i + 1
        Debug.Print "Record: " & i & ": " & rsCopySet!valName
        'Let's try one more ADO method, why not, righ?
        rsNewCoeff.AddNew


        'Add the necessary components
        rsNewCoeff!Coeff1 = 0
        rsNewCoeff!Coeff2 = 0
        rsNewCoeff!Coeff3 = 0
        rsNewCoeff!Coeff4 = 0
        rsNewCoeff!CoeffSetID = rsCoeffSetID!CoeffSetID
        rsNewCoeff!LCClassID = rsCopySet!LCClassID
        rsNewCoeff.Update
        
        rsCopySet.MoveNext
     Loop
        
    'Cleanup
    rsCopySet.Close
    rsCoeffSetID.Close
    rsNewCoeff.Close
    
    cboPollName.ListIndex = cboPollName.ListIndex
    
    cboCoeffSet.AddItem strCoeffName
    
    'Call the function to set everything to newly added Coefficient.
    cboCoeffSet.ListIndex = GetCboIndex(strCoeffName, cboCoeffSet)

    txtLCType.Text = rsLCType!Name

    Set rsCopySet = Nothing
    Set rsCoeffSetID = Nothing
    Set rsNewCoeff = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "AddCoefficient " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub UpdateValues()
  On Error GoTo ErrorHandler


    Dim i As Integer
    Dim rsPollUpdate As New ADODB.Recordset
    Dim strPollUpdate As String
    Dim strWQSelect As String
    
    Dim rsDescrip As New ADODB.Recordset
    Dim rsWQstd As New ADODB.Recordset
    
    If ValidateGridValues Then
    'Update
    
        For i = 1 To grdPollDef.Rows - 1
                            
            strPollUpdate = "SELECT * From Coefficient Where CoeffID = " & grdPollDef.TextMatrix(i, 7)
            rsPollUpdate.Open strPollUpdate, g_ADOConn, adOpenDynamic, adLockOptimistic
            
            rsPollUpdate!Coeff1 = grdPollDef.TextMatrix(i, 3)
            rsPollUpdate!Coeff2 = grdPollDef.TextMatrix(i, 4)
            rsPollUpdate!Coeff3 = grdPollDef.TextMatrix(i, 5)
            rsPollUpdate!Coeff4 = grdPollDef.TextMatrix(i, 6)
            
            rsPollUpdate.Update
            rsPollUpdate.Close
            
        Next i
        
    End If
    
    If boolDescChanged Then
        
        Dim strUpdateDescription
        strUpdateDescription = "SELECT Description from CoefficientSet Where Name like '" & cboCoeffSet.Text & "'"
    
        rsDescrip.Open strUpdateDescription, g_ADOConn, adOpenDynamic, adLockOptimistic
        
        If Len(txtCoeffSetDesc.Text) = 0 Then
            rsDescrip!Description = ""
        Else
            rsDescrip!Description = txtCoeffSetDesc.Text
        End If
        
        rsDescrip.Update
        rsDescrip.Close
        
    End If
    
    For i = 1 To grdWQStd.Rows - 1
        strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.TextMatrix(i, 4)
        
        rsWQstd.Open strWQSelect, g_ADOConn, adOpenDynamic, adLockOptimistic
        
        rsWQstd!Threshold = grdWQStd.TextMatrix(i, 3)
        rsWQstd.Update
        rsWQstd.Close

    Next i
    
    Set rsDescrip = Nothing
    Set rsPollUpdate = Nothing
    Set rsWQstd = Nothing


  Exit Sub
ErrorHandler:
  HandleError False, "UpdateValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UpdateCoeffSet(adoRSCoeff As ADODB.Recordset, strCoeffName As String, strFileName As String)
  On Error GoTo ErrorHandler


'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
'garnered above during a cbo click event.  Once that's done, we'll add a series of
'coefficients for the landclass based on the incoming textfile...strFileName

    
    Dim rsCopySet As ADODB.Recordset        'First RS
    Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    Dim rsNewCoeff As New ADODB.Recordset   'New Coefficient set
    
    Dim strNewLcType As String              'CmdString for inserting new coefficientset
    Dim strDefault As String                '
    Dim strNewCoeff As String
    Dim strNewCoeffID As String             'Holder for the CoefficientSetID
    Dim strInsertNewCoeff As String         'Putting the newly created coefficients in their table
    Dim intCoeffSetID As Integer
    Dim i As Integer
    
    'Textfile related material
    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
    Dim strLine As String
    Dim strValue As Integer
    Dim intLine As Integer
    
    adoRSCoeff.ReQuery
    
    strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & _
            Replace(strCoeffName, "'", "''") & "'," & _
            m_intPollID & "," & _
            adoRSCoeff!LCTypeID & ")"
    
    'First need to add the coefficient set to that table
    g_ADOConn.Execute strNewLcType, adAffectCurrent
    
    'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & _
        "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
        
    rsCoeffSetID.Open strNewCoeffID, g_ADOConn, adOpenKeyset, adLockReadOnly
    intCoeffSetID = rsCoeffSetID!CoeffSetID
    
    'Now turn attention to the TextFile...to get the users coefficient values
    Set fl = fso.OpenTextFile(strFileName, ForReading, True, TristateFalse)
    intLine = 0
    
    'Now loopy loo to populate values.
    Dim strNewCoeff1 As String
    strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    rsNewCoeff.Open strNewCoeff1, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'adoRSCoeff.ReQuery
    Debug.Print adoRSCoeff.RecordCount
    
    adoRSCoeff.ReQuery
    adoRSCoeff.MoveFirst
        
    i = 0
    
    grdPollDef.Clear
    grdPollDef.Rows = adoRSCoeff.RecordCount + 1
        
    Do While Not fl.AtEndOfStream
       
        strLine = fl.ReadLine
        'Value exits??
        strValue = Split(strLine, ",")(0)
        adoRSCoeff.MoveFirst
        
        
        For i = 0 To adoRSCoeff.RecordCount - 1
            Debug.Print
            
            If adoRSCoeff!Value = strValue Then
            
                'Let's try one more ADO method
                rsNewCoeff.AddNew
                
                'Add the necessary components
                rsNewCoeff!Coeff1 = Split(strLine, ",")(1)
                rsNewCoeff!Coeff2 = Split(strLine, ",")(2)
                rsNewCoeff!Coeff3 = Split(strLine, ",")(3)
                rsNewCoeff!Coeff4 = Split(strLine, ",")(4)
                rsNewCoeff!CoeffSetID = rsCoeffSetID!CoeffSetID
                rsNewCoeff!LCClassID = adoRSCoeff!LCClassID
                rsNewCoeff.Update
                
                With grdPollDef
                    .TextMatrix(i, 1) = strValue
                    .TextMatrix(i, 2) = adoRSCoeff!Name
                    .TextMatrix(i, 3) = Split(strLine, ",")(2)
                    .TextMatrix(i, 4) = Split(strLine, ",")(2)
                    .TextMatrix(i, 5) = Split(strLine, ",")(2)
                    .TextMatrix(i, 6) = Split(strLine, ",")(2)
                    .TextMatrix(i, 7) = rsCoeffSetID!CoeffSetID
                    .TextMatrix(i, 8) = rsNewCoeff!coeffID
                End With
            End If
            adoRSCoeff.MoveNext
        Next i
        
    Loop
    
    Unload frmImportCoeffSet
    cboCoeffSet.Clear
    InitComboBox cboCoeffSet, "CoefficientSet"
    cboCoeffSet.ListIndex = modUtil.GetCboIndex(strCoeffName, cboCoeffSet)
    
    


  Exit Sub
ErrorHandler:
  HandleError True, "UpdateCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub CopyCoefficient(strNewCoeffName As String, strCoeffSet As String)
  On Error GoTo ErrorHandler


'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
'from the set being copied
    Dim strCopySet As String                'The Recordset of existing coefficients being copied
    Dim strNewLcType As String              'CmdString for inserting new coefficientset               '
    Dim strNewCoeff As String
    Dim strNewCoeffID As String             'Holder for the CoefficientSetID
    Dim intCoeffSetID As Integer
    Dim i As Integer
    
    Dim rsCopySet As New ADODB.Recordset        'First RS
    Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    Dim rsNewCoeff As New ADODB.Recordset
        
    strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & _
        "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
        
    rsCopySet.Open strCopySet, g_ADOConn, adOpenKeyset
     
    
    'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
    strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & _
            Replace(strNewCoeffName, "'", "''") & "'," & _
            rsCopySet!POLLID & "," & _
            rsCopySet!LCTypeID & ")"
    
    'First need to add the coefficient set to that table
    g_ADOConn.Execute strNewLcType, adAffectCurrent
    
    'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & _
        "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"

    rsCoeffSetID.Open strNewCoeffID, g_ADOConn, adOpenKeyset, adLockPessimistic
    intCoeffSetID = rsCoeffSetID!CoeffSetID
    
    'Now loopy loo to populate values.
    Dim strNewCoeff1 As String
    strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    rsNewCoeff.Open strNewCoeff1, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    grdPollDef.Clear
    grdPollDef.Rows = rsCopySet.RecordCount + 1

    rsCopySet.MoveFirst
    
    'Actually add the records to the new set
    For i = 1 To rsCopySet.RecordCount

        'Let's try one more ADO method, why not, righ?
        rsNewCoeff.AddNew

        'Add the necessary components
        rsNewCoeff!Coeff1 = rsCopySet!Coeff1
        rsNewCoeff!Coeff2 = rsCopySet!Coeff2
        rsNewCoeff!Coeff3 = rsCopySet!Coeff3
        rsNewCoeff!Coeff4 = rsCopySet!Coeff4
        rsNewCoeff!CoeffSetID = intCoeffSetID
        rsNewCoeff!LCClassID = rsCopySet!LCClassID
    
        rsNewCoeff.Update
        
        rsCopySet.MoveNext
   
    Next i
    
    
    'Set up everything to look good
    cboPollName_Click
    cboCoeffSet.ListIndex = GetCboIndex(strNewCoeffName, cboCoeffSet)
    Unload frmImportCoeffSet


    'Cleanup
    rsCopySet.Close
    rsCoeffSetID.Close
    rsNewCoeff.Close

    Set rsCopySet = Nothing
    Set rsCoeffSetID = Nothing
    Set rsNewCoeff = Nothing

        


  Exit Sub
ErrorHandler:
  HandleError True, "CopyCoefficient " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CmdSaveEnabled()
  On Error GoTo ErrorHandler

    If boolChanged Or boolDescChanged Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If


  Exit Sub
ErrorHandler:
  HandleError False, "CmdSaveEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Form hooker
Public Sub init(ByVal pApp As IApplication)
  On Error GoTo ErrorHandler

    Set m_App = pApp

  Exit Sub
ErrorHandler:
  HandleError True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
