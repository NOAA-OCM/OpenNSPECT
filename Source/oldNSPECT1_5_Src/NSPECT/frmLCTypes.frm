VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLCTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Land Cover Types"
   ClientHeight    =   8490
   ClientLeft      =   4020
   ClientTop       =   1065
   ClientWidth     =   9270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3495
      Top             =   7995
   End
   Begin VB.CheckBox chkWWL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3105
      TabIndex        =   13
      ToolTipText     =   "Check if landuse is water or wetland"
      Top             =   7935
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cboWWL 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmLCTypes.frx":0000
      Left            =   5550
      List            =   "frmLCTypes.frx":000A
      TabIndex        =   12
      Top             =   8025
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtActiveCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   4020
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   8055
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLCClasses 
      Height          =   6675
      Left            =   105
      TabIndex        =   10
      ToolTipText     =   "Right click to add, delete, or insert a row"
      Top             =   1140
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   11774
      _Version        =   393216
      Cols            =   11
      BackColorSel    =   12648447
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   2415
      Top             =   7995
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore Defaults"
      Enabled         =   0   'False
      Height          =   375
      Left            =   375
      TabIndex        =   6
      Top             =   7905
      Width           =   1545
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   7890
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6615
      TabIndex        =   4
      Top             =   7890
      Width           =   975
   End
   Begin VB.TextBox txtLCTypeDesc 
      Height          =   285
      Left            =   1860
      TabIndex        =   2
      Top             =   420
      Width           =   5430
   End
   Begin VB.ComboBox cboLCType 
      Height          =   315
      ItemData        =   "frmLCTypes.frx":0014
      Left            =   1860
      List            =   "frmLCTypes.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUSLE"
      Height          =   285
      Index           =   2
      Left            =   6765
      TabIndex        =   9
      Top             =   810
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Classification"
      Height          =   285
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   810
      Width           =   3405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SCS Curve Numbers"
      Height          =   285
      Index           =   0
      Left            =   3540
      TabIndex        =   7
      Top             =   810
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Index           =   6
      Left            =   315
      TabIndex        =   3
      Top             =   405
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Land Cover Type:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   1
      Top             =   75
      Width           =   1365
   End
   Begin VB.Menu mnuLCTypes 
      Caption         =   "&Options"
      Begin VB.Menu mnuNewLCType 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuDelLCType 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu mnuImpLCType 
         Caption         =   "&Import..."
      End
      Begin VB.Menu mnuExpLCType 
         Caption         =   "&Export..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAppend 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuInsertRow 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuLCHelp 
         Caption         =   "Land Cover Types..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmLCTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmLCTypes
' *************************************************************************************
' *  Description: Form for browsing LC Types
' *  within NSPECT
' *
' *  Called By:  NSPECT clsLCType
' *************************************************************************************
' *  Subs:
' *     ExportLandCover(FileName as String) - creates text file of current classes
' *         Called By: mnuExpLCType
' *
' *  Misc:  Uses an invisible menu called mnuPopUp for right click events on the FlexGrid
' *************************************************************************************
Option Explicit

Private m_App As IApplication
Private rsLCTypeCboClick As ADODB.Recordset   'Recordset
Private rsCCAPDefault As ADODB.Recordset      'CCAP default recordset
Private m_intRow As Integer                   'Current Row
Private m_intCol As Integer                   'Current Col.
Private m_intLCTypeID As Long                 'LCTypeID#
Private m_intCount As Integer                 'Number of rows in old GRID
Private m_bolGridChanged As Boolean           'Flag for whether or not grid values have changed
Private m_bolSaved As Boolean                 'Flag for saved/not saved changes
Private m_bolFirstLoad As Boolean             'Is initial Load event
Private m_bolBegin As Boolean
Private m_strUndoText As String               'initial cell value used to track changes - defaults back on Esc
Private m_strUndoDescrip As String            'same but for the Description
Private m_intMouseButton As Integer           'Integer for mouse button click - added to avoid right click change cell value problem

Private clsLCClassData As New clsLCClassData  'Class that handles the data
Dim WithEvents m_adoConn As ADODB.Connection  'BeginTrans, CommitTrans Event handler
Attribute m_adoConn.VB_VarHelpID = -1

Private Sub chkWWL_Click(Index As Integer)
    
    grdLCClasses.TextMatrix(Index, 8) = chkWWL(Index).Value
    
    If Not m_bolFirstLoad Then
        'm_bolGridChanged = True
        'MsgBox m_bolGridChanged
        CmdSaveEnabled
    End If
    
End Sub

Private Sub chkWWL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MsgBox "MouseDown"
    m_bolGridChanged = True
    CmdSaveEnabled
End Sub

Private Sub cmdRestore_Click()
'Restore Defaults Button - just read in NSPECT.LCCLASSDEFAULTS

On Error GoTo Errhand:
    
    Dim strCCAP As String
       
    'Check to make sure that's what they want
    intYesNo = MsgBox(strDefault, vbYesNo, strDefaultTitle)
    
    If intYesNo = vbYes Then

        Dim i As Integer
        
        Set rsCCAPDefault = New ADODB.Recordset
        
        'Selection based on combo box
        strCCAP = "SELECT * From LCCLASSDEFAULTS"
        
        rsCCAPDefault.Open strCCAP, modUtil.g_ADOConn, adOpenKeyset, adLockOptimistic
        rsCCAPDefault.MoveFirst
        
        grdLCClasses.Rows = rsCCAPDefault.RecordCount + 1
        grdLCClasses.Refresh
        m_intCount = grdLCClasses.Rows
        
        'Clear out the old dataset - again, column 10 contains the LCClassID
        For i = 1 To rsCCAPDefault.RecordCount 'grdLCClasses.Rows - 1
            
            grdLCClasses.TextMatrix(i, 1) = rsCCAPDefault!Value
            grdLCClasses.TextMatrix(i, 2) = rsCCAPDefault!Name
            grdLCClasses.TextMatrix(i, 3) = rsCCAPDefault![CN-A]
            grdLCClasses.TextMatrix(i, 4) = rsCCAPDefault![CN-B]
            grdLCClasses.TextMatrix(i, 5) = rsCCAPDefault![CN-C]
            grdLCClasses.TextMatrix(i, 6) = rsCCAPDefault![CN-D]
            grdLCClasses.TextMatrix(i, 7) = rsCCAPDefault!CoverFactor
            grdLCClasses.TextMatrix(i, 8) = rsCCAPDefault!W_WL
            
            rsCCAPDefault.MoveNext
        
        Next i
        
        ClearCheckBoxes True, m_intCount
        CreateCheckBoxes True
        
        rsCCAPDefault.Close
        Set rsCCAPDefault = Nothing
        
        m_bolGridChanged = True
        CmdSaveEnabled
        
    Else
        
        Exit Sub
    
    End If
    
    Exit Sub

Errhand:
    MsgBox "There was an error loading the default CCAP data.", vbCritical, "Error Loading Data"
    
    
End Sub

Private Sub cmdSave_Click()

On Error GoTo ErrHandler:
    txtActiveCell.Visible = False
    
    If ValidateGridValues Then
        UpdateValues
        g_ADOConn.CommitTrans
        
        MsgBox "Data saved successfully.", vbOKOnly, "Data Saved Successfully"
        
        'Reset the flags
        m_bolGridChanged = False
        m_bolSaved = True
        
        Set m_adoConn = Nothing
        
        Unload Me
    End If
Exit Sub
ErrHandler:
    
    If Err.Number = -2147221504 Then
        MsgBox "The data values entered exceed the allowable precision of the database." & vbNewLine & _
        "Data must not contain more than 4 values to the right of the decimal place." & vbNewLine & _
        "Please correct your inputs before saving.", vbInformation, "Precision Error"
        Exit Sub
    End If
    
    MsgBox "There was an error saving changes.", vbCritical, "Error Saving Changes"
    
    
    
    Exit Sub

End Sub



Private Sub Form_Click()
'If user clicks the form, get rid of the txtActiveCell
    
    If txtActiveCell.Visible = True Then
        grdLCClasses.Text = txtActiveCell.Text
        txtActiveCell.Visible = False
    End If
    
    
End Sub

Private Sub Form_Load()

    Set m_adoConn = g_ADOConn
    
    'Set the flags
    m_bolSaved = False              'We haven't saved
    m_bolGridChanged = False        'Nothing's changed
    m_bolFirstLoad = True           'It's the first load
    
    'Initialize the Grid and populate the combobox
    modUtil.InitComboBox cboLCType, "LCTYPE"
    
End Sub


Private Sub cboLCType_Click()

    'On the cbo Click to change to a new LandClassType, check if there's been changes, prompt to save
    If m_bolGridChanged And m_bolBegin Then
        
        intYesNo = MsgBox(strYesNo, vbYesNo, strYesNoTitle)
            'Since we're changing records for LCTypes...good time to save changes, ergo CommitTrans
            If intYesNo = vbYes Then
                UpdateValues
                g_ADOConn.CommitTrans
                m_bolGridChanged = False
                CmdSaveEnabled
            ElseIf intYesNo = vbNo Then
                g_ADOConn.RollbackTrans
                m_bolGridChanged = False
                CmdSaveEnabled
            End If
            
    Else
        
        m_intCount = grdLCClasses.Rows
        CheckCCAPDefault cboLCType.Text
        txtActiveCell.Visible = False
        
        'original
        Dim strSQLLCType As String
        Set rsLCTypeCboClick = New ADODB.Recordset
    
        Dim strSQLLCClass As String
    
        'Selection based on combo box
        strSQLLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"
    
        With rsLCTypeCboClick
            .CursorLocation = adUseClient
            .Open strSQLLCType, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
        End With
    
        If rsLCTypeCboClick.RecordCount > 0 Then
            
            txtLCTypeDesc.Text = rsLCTypeCboClick!Description
                        
            strSQLLCClass = "SELECT LCCLASS.Value, LCCLASS.Name, LCCLASS.[CN-A], LCCLASS.[CN-B]," & _
            " LCCLASS.[CN-C], LCCLASS.[CN-D], LCCLASS.CoverFactor, LCCLASS.W_WL, LCCLASS.LCTYPEID, LCCLASS.LCCLASSID FROM LCCLASS WHERE" & _
            " LCTYPEID = " & rsLCTypeCboClick!LCTypeID & " ORDER BY LCCLass.Value"
            
            With clsLCClassData
                .SQL = strSQLLCClass
            End With
                                 
            With grdLCClasses
                Set .Recordset = clsLCClassData.Recordset
                modUtil.InitLCClassesGrid grdLCClasses
                .Refresh
            End With
            
            clsLCClassData.Recordset.MoveFirst
            'Get the LCTypeID for use in updates, deletes, etc
            m_intLCTypeID = clsLCClassData.Recordset!LCTypeID
                
            If Not m_bolBegin Then
                g_ADOConn.BeginTrans
            End If
        Else
        
            MsgBox "Warning: There are no records remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        
        End If
    
    End If
    
    If Not m_bolBegin Then
        g_ADOConn.BeginTrans
    End If
    
    Timer1.Interval = 10
    Timer1.Enabled = True

End Sub


Private Sub mnuLCHelp_Click()

    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "land_cover.htm"
    
End Sub

Private Sub Timer1_Timer()
   
    Timer1.Enabled = False
    If m_bolFirstLoad Then
        CreateCheckBoxes True
        m_bolFirstLoad = False
    Else
        ClearCheckBoxes True, m_intCount
        CreateCheckBoxes True
    End If
    
End Sub

Private Sub cmdQuit_Click()
   
    If Not m_bolSaved And m_bolGridChanged Then
        
        intYesNo = MsgBox("Do you want to save changes made to " & cboLCType.Text & "?", vbYesNoCancel + vbExclamation, "N-SPECT")
                
            If intYesNo = vbYes Then
            
                If ValidateGridValues Then
                    UpdateValues
                    g_ADOConn.CommitTrans
                    MsgBox "Data saved successfully.", vbOKOnly, "Save Successful"
                    m_bolGridChanged = False
                    m_bolSaved = True
                    Unload Me
                End If
            
            ElseIf intYesNo = vbNo Then
                
                g_ADOConn.RollbackTrans
                Unload Me
                
            Else
                Exit Sub
            
            End If
    Else
        g_ADOConn.CommitTrans
        Unload Me
    End If
    
End Sub

Private Sub CmdSaveEnabled()

    cmdSave.Enabled = m_bolGridChanged

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_adoConn = Nothing
End Sub
Private Sub m_adoConn_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    m_bolBegin = True
End Sub
Private Sub m_adoConn_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    m_bolBegin = False
End Sub

Private Sub m_adoConn_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    m_bolBegin = False
End Sub

'''''''''''''''''''''''''''''''''''''''
'OPTIONS Menu
'''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''
'NEW Menu
'''''''''''''''''''''''''''
Private Sub mnuNewLCType_Click()
   frmNewLCType.Show vbModal, Me
End Sub
   
'''''''''''''''''''''''''''
'IMPORT Menu
'''''''''''''''''''''''''''
Private Sub mnuImpLCType_Click()
   frmImportLCType.Show vbModal, Me
End Sub

'''''''''''''''''''''''''''
'DELETE Menu
'''''''''''''''''''''''''''
Private Sub mnuDelLCType_Click()
   
   Dim intAns As Integer
   intAns = MsgBox("Are you sure you want to delete the land cover type '" & cboLCType.List(cboLCType.ListIndex) & "' and all associated Coefficient Sets?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   
   If intAns = vbYes Then
   
    Dim strLCTypeDelete As String
    Dim strLCClassDelete As String
    
    Dim rsLCTypeDelete As ADODB.Recordset
    Dim rsLCClassDelete As ADODB.Recordset
    
    strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"
    
    Set rsLCTypeDelete = New ADODB.Recordset
    
    rsLCTypeDelete.CursorLocation = adUseClient
    rsLCTypeDelete.Open strLCTypeDelete, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
       
    strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & rsLCTypeDelete!LCTypeID
       
    If Not (cboLCType.Text = "") Then
    
    'code to handle response
             
         modUtil.g_ADOConn.Execute strLCClassDelete
        
        'Set up a delete rs and get rid of it
        rsLCTypeDelete.Delete adAffectCurrent
        rsLCTypeDelete.Update
        
        MsgBox cboLCType.List(cboLCType.ListIndex) & " deleted.", vbOKOnly, "Record Deleted"
        
        cboLCType.Clear
        m_bolGridChanged = False
        modUtil.InitComboBox cboLCType, "LCType"
        frmLCTypes.Refresh
                            
    Else
        MsgBox "Please select a Land class", vbCritical, "No Land Class Selected"
    End If
    
    
    
    ElseIf intAns = vbNo Then
        m_bolGridChanged = False
        Exit Sub
    End If
    

End Sub
   
''''''''''''''''''''''''''
'EXPORT Menu
''''''''''''''''''''''''''
Private Sub mnuExpLCType_Click()
  Dim intAns As Integer
   
   'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = Replace(MSG1, "<name>", "Land Cover Type")
     .DialogTitle = Replace(MSG3, "<name>", "Land Cover Type")
     .FilterIndex = 1
     .DefaultExt = ".txt"
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
     .ShowSave
   End With
   If Len(dlgCMD1.FileName) > 0 Then
      'Export land cover type to file - dlgCMD1.FileName
      ExportLandCover dlgCMD1.FileName
   End If

   
End Sub

'''***************************************
'''mnuPopup Functionality - Andrew
'''***************************************
'Add row...
Private Sub mnuAppend_Click()

    clsLCClassData.LCTypeID = getLCTypeID
    clsLCClassData.AddNew
    
    'add row to end of  grid
    With grdLCClasses
       .Rows = .Rows + 1
       .row = .Rows - 1
       .TextMatrix(.row, 0) = ""
       .TextMatrix(.row, 1) = "0"
       .TextMatrix(.row, 2) = "Landclass" & .row
       .TextMatrix(.row, 3) = "0"
       .TextMatrix(.row, 4) = "0"
       .TextMatrix(.row, 5) = "0"
       .TextMatrix(.row, 6) = "0"
       .TextMatrix(.row, 7) = "0"
       .TextMatrix(.row, 8) = "0"
       .TextMatrix(.row, 9) = clsLCClassData.LCTypeID
       .TextMatrix(.row, 10) = g_intLCClassid
    
    End With
    'IsCellVisible
    
    m_intCount = grdLCClasses.Rows
    
    CreateCheckBoxes False, grdLCClasses.row
    
    m_bolGridChanged = True
    CmdSaveEnabled
   
End Sub

Private Sub mnuDeleteRow_Click()
'delete current row

    Dim R%, C%, row%, col%
    Dim lngLCClassID As Long
    
    With grdLCClasses
    
      lngLCClassID = .TextMatrix(m_intRow, 10)
      
      If .Rows > .FixedRows Then        'make sure we don't del header Rows
          For col% = 1 To .Cols - 1
              If ((Trim(.TextMatrix(.row, col%)) > "" And col% = 2) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% = 1) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% >= 3)) Then 'data?
                  C% = 1
                  Exit For
              End If
          Next col%
          If C% Then
              R% = MsgBox("There is data in Row" + Str$(.row) + " ! Delete anyway?", vbYesNo, "Delete Row!")
          End If
          If C% = 0 Or R% = vbYes Then        'no exist. data or YES
              If .row = .Rows - 1 Then  'last row?
                  .row = .row - 1       'move active cell
              Else
                  For row% = .row To .Rows - 2 'move data up 1 row
                      For col% = 1 To .Cols - 1
                          .TextMatrix(row%, col%) = .TextMatrix(row% + 1, col%)
                      Next col%
                  Next row%
              End If
               .Rows = .Rows - 1 'del last row
               clsLCClassData.Load lngLCClassID
               clsLCClassData.Delete
          End If
      End If
   End With
   
   m_intCount = grdLCClasses.Rows
   ClearCheckBoxes True, m_intCount + 1
   CreateCheckBoxes True
   
   m_bolGridChanged = True       'reset
   CmdSaveEnabled

End Sub

Private Sub mnuInsertRow_Click()
'Insert row above current row in grdLCClasses- Thanks, Andrew
        
    'Get a hold of LCTYPEID, must have it to insert new records
    clsLCClassData.LCTypeID = m_intLCTypeID
    clsLCClassData.AddNew
    
    Dim R%, row%, col%
    With grdLCClasses
      If .row < .FixedRows Then        'make sure we don't insert above header Rows
         mnuAppend_Click
      Else
         R% = .row
         .Rows = .Rows + 1             'add a row
         
         For row% = .Rows - 1 To R% + 1 Step -1 'move data dn 1 row
             For col% = 1 To .Cols - 1
                 .TextMatrix(row%, col%) = .TextMatrix(row% - 1, col%)
             Next col%
         Next row%
         For col% = 1 To .Cols - 1       ' clear all cells in this row
            If (col% = 2) Then
               .TextMatrix(R%, col%) = ""
            Else
               .TextMatrix(R%, col%) = "0"
            End If
         Next col%
         .TextMatrix(R%, 9) = clsLCClassData.LCTypeID
         .TextMatrix(R%, 10) = g_intLCClassid
         
         
     End If
   End With
    
   txtActiveCell.Visible = False
   m_intCount = grdLCClasses.Rows
   
   ClearCheckBoxes True, m_intCount - 1
   CreateCheckBoxes True
   
   m_bolGridChanged = True                'reset
   CmdSaveEnabled
   
End Sub

Private Sub grdLCClasses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_intRow = grdLCClasses.MouseRow
    m_intCol = grdLCClasses.MouseCol

    m_intMouseButton = Button
    
    If Button = 2 Then
        
        If m_intCol = 0 And cboLCType.Text <> "CCAP" Then
            With grdLCClasses
                .row = m_intRow
                .col = m_intCol
            End With
            modUtil.HighlightGridRow grdLCClasses, m_intRow
            PopupMenu mnuPopUp
        End If
        
    End If

End Sub

Private Sub grdLcClasses_Click()
 
        m_intRow = grdLCClasses.MouseRow
        m_intCol = grdLCClasses.MouseCol
        
        txtActiveCell.Visible = False
        
        If m_intMouseButton = 1 Then
        
            If (m_intCol >= 1 And m_intCol < 8) And (m_intRow >= 1) Then
                
                With txtActiveCell
                
                     .Move grdLCClasses.Left + grdLCClasses.CellLeft, _
                     grdLCClasses.Top + grdLCClasses.CellTop, _
                     grdLCClasses.CellWidth - 30, _
                     grdLCClasses.CellHeight - 75
                     
                    If m_intRow <> 0 Then
                        .Visible = True
                        m_strUndoText = grdLCClasses.TextMatrix(m_intRow, m_intCol)
                        .Text = grdLCClasses.TextMatrix(m_intRow, m_intCol)
                        .SetFocus
                        .SelLength = Len(.Text)
                        
                        If m_intCol >= 3 And m_intCol <= 7 Then
                            txtActiveCell.Alignment = 0
                        Else
                            If IsNumeric(m_strUndoText) Then
                                txtActiveCell.Alignment = 1
                            Else
                                txtActiveCell.Alignment = 0
                            End If
                        End If
                    End If
                End With
                
            ElseIf m_intCol = 0 Then
                
                txtActiveCell.Visible = False
            
            End If
        Else
            Exit Sub
        End If
End Sub

Private Sub txtActiveCell_Change()
    
    grdLCClasses.Text = txtActiveCell.Text
    
End Sub


Private Sub txtActiveCell_GotFocus()
    
    m_strUndoText = txtActiveCell.Text
    
    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)

End Sub

Private Sub txtActiveCell_Validate(Cancel As Boolean)
    
If Not m_bolGridChanged Then
    
    If txtActiveCell.Text <> m_strUndoText Then
        m_bolGridChanged = True
        CmdSaveEnabled
    End If
    
End If

End Sub

Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)
    
'Handles some key inputs from the txtActiveCell, basically provides cell to cell movement around the
'GRID
    With grdLCClasses
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
                If .row > 0 Then
                    .row = .row - 1
                    KeyMoveUpdate
                Else
                    .row = .row 'if the row is already on zero, don't move cells
                End If
            Case vbKeyDown 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                If .row < grdLCClasses.Rows - 1 Then
                    .row = .row + 1
                    KeyMoveUpdate
                Else
                    .row = .row 'again, if the row is on the last row, don't move cells
                    .Text = txtActiveCell.Text
                    txtActiveCell.Visible = False
                End If
            Case vbKeyLeft 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                txtActiveCell.Visible = False
                If .col > 1 Then
                    .col = .col - 1
                    KeyMoveUpdate
                Else
                
                If .row > 1 Then
                    .col = 7
                    .row = .row - 1
                    KeyMoveUpdate
                End If
                End If
            Case vbKeyRight 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                '.Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                If .col < 7 Then
                    .col = .col + 1
                    KeyMoveUpdate
                Else
                    If .row < grdLCClasses.Rows - 1 Then
                        .col = 1
                        .row = .row + 1
                        KeyMoveUpdate
                    End If
                End If
            Case vbKeyTab
                MsgBox "WWWWWWWWWWWOOOOOO"
        End Select
    End With
End Sub

Private Sub KeyMoveUpdate()

'This guy basically replicates the functionality of the grdLCClasses_Click event and is used
'in a couple of instances for moving around the grid.
 
    m_intRow = grdLCClasses.row
    m_intCol = grdLCClasses.col

    'chkWWL.Visible = False
    txtActiveCell.Visible = False

    If (m_intCol >= 1 And m_intCol < 8) And (m_intRow >= 1) Then
        
        With txtActiveCell
        
             .Move grdLCClasses.Left + grdLCClasses.CellLeft, _
             grdLCClasses.Top + grdLCClasses.CellTop, _
             grdLCClasses.CellWidth - 30, _
             grdLCClasses.CellHeight - 75
             
            If m_intRow <> 0 Then
                
                .Visible = True
                m_strUndoText = grdLCClasses.TextMatrix(m_intRow, m_intCol)
                .Text = m_strUndoText
                
                If m_intCol >= 3 And m_intCol <= 7 Then
                    txtActiveCell.Alignment = 0
                Else
                    If IsNumeric(m_strUndoText) Then
                        txtActiveCell.Alignment = 1
                    Else
                        txtActiveCell.Alignment = 0
                    End If
                End If
                
            End If
        End With
        txtActiveCell.SetFocus
    
    ElseIf m_intCol = 0 Then
        
        txtActiveCell.Visible = False
    
    End If
    
End Sub

'''***************************************
'''Subs/Functions
'''***************************************
Private Sub ExportLandCover(strFileName As String)
'Exports your current LCType/LCClasses to text or csv.

    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
    Dim rsNew As Recordset
    Dim theLine
    
    Set fl = fso.CreateTextFile(strFileName, True)
    
    'Write the name and descript.
    With fl
        .WriteLine cboLCType.Text & "," & txtLCTypeDesc.Text
    End With
    
    Dim i As Integer
    
    'Write name of pollutant and threshold
    For i = 1 To grdLCClasses.Rows - 1
        fl.WriteLine grdLCClasses.TextMatrix(i, 1) & "," & grdLCClasses.TextMatrix(i, 2) & _
        "," & grdLCClasses.TextMatrix(i, 3) & _
        "," & grdLCClasses.TextMatrix(i, 4) & _
        "," & grdLCClasses.TextMatrix(i, 5) & _
        "," & grdLCClasses.TextMatrix(i, 6) & _
        "," & grdLCClasses.TextMatrix(i, 7) & _
        "," & grdLCClasses.TextMatrix(i, 8)
        
    Next i

    fl.Close
        
    Set fso = Nothing
    Set fl = Nothing

End Sub

Private Function ValidateGridValues() As Boolean

'Need to validate each grid value before saving.  Essentially we take it a row at a time,
'then rifle through each column of each row.  Case Select tests each each x,y value depending
'on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

'Returns: True or False

    Dim varActive As Variant            'txtActiveCell value
    Dim varColumn2Value As Variant      'Value of Column 2 ([VALUE]) - have to check for unique
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    For i = 1 To grdLCClasses.Rows - 1
        
        For j = 1 To 7
    
        varActive = grdLCClasses.TextMatrix(i, j)
        
        Select Case j
        
            Case 1
                If Not IsNumeric(varActive) Then
                    ErrorGenerator Err1, i, j
                Else
                    For k = 1 To grdLCClasses.Rows - 1
                        
                        varColumn2Value = grdLCClasses.TextMatrix(k, 1)
                        If k <> i Then 'Don't want to compare value to itself
                            If varColumn2Value = grdLCClasses.TextMatrix(i, 1) Then
                                ErrorGenerator Err2, i, j
                                grdLCClasses.col = j
                                grdLCClasses.row = i
                                ValidateGridValues = False
                                KeyMoveUpdate
                                Exit Function
                            End If
                        End If
                    Next k
                End If
                    
        
            Case 2
                If IsNumeric(varActive) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
        
            Case 3
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
                
            Case 4
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
                
            Case 5
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
            
            Case 6
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
            
            Case 7
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 5) Then
                    ErrorGenerator Err3, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
            End If
            End Select
            Next j
    Next i
        
    ValidateGridValues = True
        
End Function

Private Function getLCTypeID() As Long

    getLCTypeID = m_intLCTypeID 'grdLCClasses.TextMatrix(1, 9)

End Function


Private Sub CheckCCAPDefault(strName As String)
    
    If strName = "CCAP" Then
        cmdRestore.Enabled = True
        mnuDelLCType.Enabled = False
        mnuPopUp.Enabled = False
        grdLCClasses.ToolTipText = ""
    Else
        cmdRestore.Enabled = False
        mnuDelLCType.Enabled = True
        mnuPopUp.Enabled = True
        grdLCClasses.ToolTipText = "Right click to add, delete, or insert a row"
    End If
    
End Sub

Private Sub UpdateValues()
    
    Dim i As Integer
    
    For i = 1 To grdLCClasses.Rows - 1
                
        With clsLCClassData
            
            .Load grdLCClasses.TextMatrix(i, 10)
            .Value = grdLCClasses.TextMatrix(i, 1)
            .Name = grdLCClasses.TextMatrix(i, 2)
            .CNA = grdLCClasses.TextMatrix(i, 3)
            .CNB = grdLCClasses.TextMatrix(i, 4)
            .CNC = grdLCClasses.TextMatrix(i, 5)
            .CND = grdLCClasses.TextMatrix(i, 6)
            .CoverFactor = grdLCClasses.TextMatrix(i, 7)
            .W_WL = grdLCClasses.TextMatrix(i, 8)
            .LCTypeID = grdLCClasses.TextMatrix(i, 9)
            .LCClassID = grdLCClasses.TextMatrix(i, 10)
            
            .SaveChanges
            
        End With
    Next i
            
End Sub

Private Sub AddDefaultValues()
    
    Dim i As Integer
    
    For i = 1 To grdLCClasses.Rows - 1
                
        With clsLCClassData
        
            .Value = grdLCClasses.TextMatrix(i, 1)
            .Name = grdLCClasses.TextMatrix(i, 2)
            .CNA = grdLCClasses.TextMatrix(i, 3)
            .CNB = grdLCClasses.TextMatrix(i, 4)
            .CNC = grdLCClasses.TextMatrix(i, 5)
            .CND = grdLCClasses.TextMatrix(i, 6)
            .CoverFactor = grdLCClasses.TextMatrix(i, 7)
            .W_WL = grdLCClasses.TextMatrix(i, 8)
            .LCTypeID = grdLCClasses.TextMatrix(i, 9)
                        
            .AddNew
               
        End With
    Next i
    
    ClearCheckBoxes True, m_intCount
    CreateCheckBoxes True
    
            
End Sub

Private Sub CreateCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strChkName As String
    j = 1
    i = 1
    
    If booAll Then
        For i = 1 To grdLCClasses.Rows - 1
            
            grdLCClasses.row = i
            grdLCClasses.col = 8
            
            'Set the alignment to cover up current numbers
            grdLCClasses.CellAlignment = flexAlignCenterCenter
            
            k = i
            
            Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
            Load chkWWL(k)
            Set chkWWL(k).Container = frmLCTypes
            With chkWWL(k)
                .Visible = True
                .Top = grdLCClasses.Top + grdLCClasses.CellTop
                .Left = (grdLCClasses.Left) + (grdLCClasses.CellLeft) + (grdLCClasses.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                .Value = CInt(grdLCClasses.TextMatrix(i, 8))
            End With
           Call Controls.Remove("chk" & CStr(k))
        Next i
    Else
            
            With grdLCClasses
                .row = intRecNo
                .col = 8
                .CellAlignment = flexAlignCenterCenter
            End With
            
            k = intRecNo
            
            Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
            Load chkWWL(k)
            Set chkWWL(k).Container = frmLCTypes
            With chkWWL(k)
                .Visible = True
                .Top = grdLCClasses.Top + grdLCClasses.CellTop
                .Left = grdLCClasses.Left + grdLCClasses.CellLeft + (grdLCClasses.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                .Value = CInt(grdLCClasses.TextMatrix(k, 8))
            End With
           Call Controls.Remove("chk" & CStr(k))
    End If

End Sub

Private Sub ClearCheckBoxes(booAll As Boolean, intCount As Integer, Optional intRecNo As Integer)

    Dim k As Integer
    
    If booAll Then
    
        For k = 1 To chkWWL().UBound 'intCount - 1
            Unload chkWWL(k)
        Next k
    Else
        Unload chkWWL(intRecNo)
    End If
End Sub


Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub
