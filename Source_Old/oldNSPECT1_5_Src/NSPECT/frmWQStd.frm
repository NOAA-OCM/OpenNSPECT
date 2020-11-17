VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWQStd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Water Quality Standards"
   ClientHeight    =   4170
   ClientLeft      =   135
   ClientTop       =   405
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtActiveCell 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1965
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3585
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWQStd 
      Height          =   2415
      Left            =   345
      TabIndex        =   6
      Top             =   975
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   4
      BackColorSel    =   8388608
      ForeColorSel    =   -2147483640
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3570
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3705
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3570
      Width           =   975
   End
   Begin VB.ComboBox cboWQStdName 
      Height          =   315
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2400
   End
   Begin VB.TextBox txtWQStdDesc 
      Height          =   300
      Left            =   1455
      MaxLength       =   100
      TabIndex        =   1
      Top             =   510
      Width           =   4290
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   60
      Top             =   3990
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   285
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Standard Name:"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   120
      Width           =   1425
   End
   Begin VB.Menu mnuWQStd 
      Caption         =   "&Options"
      Begin VB.Menu mnuNewWQStd 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuDelWQStd 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuCopyWQStd 
         Caption         =   "Copy..."
      End
      Begin VB.Menu mnuImpWQStd 
         Caption         =   "Import..."
      End
      Begin VB.Menu mnuExpWQStd 
         Caption         =   "Export..."
      End
   End
   Begin VB.Menu mnuEditCell 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRow 
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
      Begin VB.Menu mnuWQHelp 
         Caption         =   "Water Quality Standards..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmWQStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmWQStd
' *************************************************************************************
' *  Description: Form for browsing and maintaining water quality standards
' *  within NSPECT
' *
' *  Called By:  clsWQStd
' *************************************************************************************
Option Explicit
Private rsWQStdLoad As ADODB.Recordset      'Load recordset
Private rsWQStdCboClick As ADODB.Recordset  'cbo Click event
Private rsWQStdPoll As ADODB.Recordset      'Pollutants of current standard
Private rsWQStdDelete As ADODB.Recordset    'RS to be deleted

Private m_App As IApplication
Private m_strFileName As String
Private strUndoText As String

Private m_strUndoText As String

Private intRow As Integer                   'Current Row
Private intCol As Integer                   'Current Col.

Private m_intWQRow As Integer
Private m_intWQCol As Integer

Private m_bolChange As Boolean                  'Have records changed?
Private intNumChanges As Integer            'Counter for changes

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmWQStd.frm"

Public Sub init(ByVal pApp As IApplication)
  On Error GoTo ErrorHandler

    Set m_App = pApp

  Exit Sub
ErrorHandler:
  HandleError True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cboWQStdName_Click()
  On Error GoTo ErrorHandler

    
    Dim intYesNo As Integer
    
    If m_bolChange Then
        intYesNo = MsgBox("You have made changes to the data.  Would you like to save before coninuing?", vbYesNo, "Save Changes?")
        If intYesNo = vbYes Then
            UpdateData
            m_bolChange = False
        ElseIf intYesNo = vbNo Then
            m_bolChange = False
        End If
        
    End If
        
    Dim strSQLWQStd As String
    Set rsWQStdCboClick = New ADODB.Recordset
    
    Dim strSQLWQStdPoll As String
    Set rsWQStdPoll = New ADODB.Recordset

    
    'Selection based on combo box
    strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"
    rsWQStdCboClick.CursorLocation = adUseClient
    rsWQStdCboClick.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    If rsWQStdCboClick.RecordCount > 0 Then
        
        txtWQStdDesc.Text = rsWQStdCboClick!Description
        
        strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD, POLL_WQCRITERIA.POLL_WQCRITID " & _
        "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & _
        "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick!WQCRITID
             
        rsWQStdPoll.CursorLocation = adUseClient
        rsWQStdPoll.Open strSQLWQStdPoll, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
        Set grdWQStd.Recordset = rsWQStdPoll
               
        modUtil.InitWQStdGrid grdWQStd
               
        'Clean it
        Set rsWQStdPoll = Nothing
        Set rsWQStdCboClick = Nothing
        
    Else
        
        MsgBox "Warning: There are no water quality standards remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        Set rsWQStdCboClick = Nothing
    
    End If



  Exit Sub
ErrorHandler:
  HandleError True, "cboWQStdName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdSave_Click()

On Error GoTo ErrHandler:
    If ValidateData Then
        UpdateData
        MsgBox "Data saved successfully.", vbInformation, "Data Saved"
    End If

Exit Sub
ErrHandler:
    MsgBox "Error updating Water Quality Standards: " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Click()
  On Error GoTo ErrorHandler


    txtActiveCell.Visible = False
    


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

  
   'Initialize the grid and then populate
    
    modUtil.InitComboBox cboWQStdName, "WQCRITERIA"
    modUtil.InitWQStdGrid grdWQStd
    'Set some values
    m_bolChange = False
    intNumChanges = 0


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    Unload frmWQStd


  Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub grdWQStd_Click()
  On Error GoTo ErrorHandler

    
    txtActiveCell.Visible = False
   
    intRow = grdWQStd.row
    intCol = grdWQStd.col

    If intCol = 2 Then
        
        With txtActiveCell
        
             .Move grdWQStd.Left + grdWQStd.CellLeft, _
             grdWQStd.Top + grdWQStd.CellTop, _
             grdWQStd.CellWidth - 30, _
             grdWQStd.CellHeight - 75
             
            If intRow <> 0 Then
                .Visible = True
                strUndoText = grdWQStd.TextMatrix(intRow, intCol)
                .Text = grdWQStd.TextMatrix(intRow, intCol)
                .SetFocus
            End If
            
        End With

    End If



  Exit Sub
ErrorHandler:
  HandleError True, "grdWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'Options menu option processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewWQStd_Click()
  On Error GoTo ErrorHandler

   frmAddWQStd.Show vbModal, Me


  Exit Sub
ErrorHandler:
  HandleError True, "mnuNewWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuCopyWQStd_Click()
  On Error GoTo ErrorHandler


   frmCopyWQStd.Show vbModal, Me



  Exit Sub
ErrorHandler:
  HandleError True, "mnuCopyWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuDelWQStd_Click()

On Error GoTo ErrHandler:
   Dim intAns As Integer
   intAns = MsgBox("Are you sure you want to delete the Water Quality Standard '" & cboWQStdName.List(cboWQStdName.ListIndex) & "'?", vbYesNo + vbDefaultButton2, "Confirm Delete")
   'code to handle response
   
   'WQ Recordset
   Dim strWQStdDelete As String
   Dim strWQPollDelete As String
   
   strWQStdDelete = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboWQStdName.Text & "'"
   
   Set rsWQStdDelete = New ADODB.Recordset
   
   rsWQStdDelete.CursorLocation = adUseClient
   rsWQStdDelete.Open strWQStdDelete, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
      
   strWQPollDelete = "Delete * FROM POLL_WQCRITERIA WHERE WQCRITID =" & rsWQStdDelete!WQCRITID
      
   If Not (cboWQStdName.Text = "") Then
   'code to handle response
        If intAns = vbYes Then
            
            'Delete the WaterQuality Standard
            rsWQStdDelete.Delete adAffectCurrent
            rsWQStdDelete.Update
        
            'modUtil.g_ADOConn.Execute strWQPollDelete
            
            MsgBox cboWQStdName.List(cboWQStdName.ListIndex) & " deleted.", vbOKOnly, "Record Deleted"
            
            cboWQStdName.Clear
            modUtil.InitComboBox cboWQStdName, "WQCRITERIA"
            frmWQStd.Refresh
                  
        ElseIf intAns = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "Please select a water quality standard", vbCritical, "No Standard Selected"
    End If
    
    'Cleanup
    rsWQStdDelete.Close
    Set rsWQStdDelete = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "An Error occurred during deletion." & "  " & Err.Number & ": " & Err.Description, vbCritical, "Error"

End Sub

'Import Menu
Private Sub mnuImpWQStd_Click()
  On Error GoTo ErrorHandler

   frmImportWQStd.Show vbModal, Me

  Exit Sub
ErrorHandler:
  HandleError True, "mnuImpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Export Menu
Private Sub mnuExpWQStd_Click()
  On Error GoTo ErrorHandler

   Dim intAns As Integer
   
   'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = Replace(MSG1, "<name>", "Water Quality Standard")
     .DialogTitle = Replace(MSG3, "<name>", "Water Quality Standard")
     .FilterIndex = 1
     .DefaultExt = ".txt"
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
     .ShowSave
   End With
   If Len(dlgCMD1.FileName) > 0 Then
      'Export Water Quality Standard to file - dlgCMD1.FileName
      ExportStandard dlgCMD1.FileName
   End If



  Exit Sub
ErrorHandler:
  HandleError True, "mnuExpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuWQHelp_Click()
    
    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "wq_stnds.htm"
    
End Sub

Private Sub txtActiveCell_Change()
  On Error GoTo ErrorHandler

'See the grd text to the text box
    grdWQStd.Text = txtActiveCell.Text
    m_bolChange = True
      


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'Exports your current standard and pollutants to text or csv.
Private Sub ExportStandard(strFileName As String)
  On Error GoTo ErrorHandler


    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
     
    Dim rsNew As Recordset
    Dim theLine
    
    Set fl = fso.CreateTextFile(strFileName, True)
    
    'Write the name and descript.
    With fl
        .WriteLine cboWQStdName.Text & "," & txtWQStdDesc
    End With
    
    Dim i As Integer
    
    'Write name of pollutant and threshold
    For i = 1 To grdWQStd.Rows - 1
        fl.WriteLine grdWQStd.TextMatrix(i, 1) & "," & grdWQStd.TextMatrix(i, 2)
    Next i

    fl.Close
        
    Set fso = Nothing
    Set fl = Nothing



  Exit Sub
ErrorHandler:
  HandleError False, "ExportStandard " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_GotFocus()
  On Error GoTo ErrorHandler

    
    strUndoText = txtActiveCell.Text
    
    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)
    


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler

'Handles some key inputs
    With grdWQStd

        Select Case KeyCode
            Case vbKeyEscape 'if the user pressed escape, then get out without changing
                '.Text = UndoText
                txtActiveCell.Visible = False
                .SetFocus
            Case 13 'if the user presses enter, get out of the textbox
                txtActiveCell.Visible = False
                .SetFocus
            Case vbKeyUp 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                .SetFocus
                If .row > 0 Then
                    .row = .row - 1
                Else
                    .row = .row 'if the row is already on zero, don't move cells
                End If
                KeyMoveUpdate
            Case vbKeyDown 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
                .Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                .SetFocus
                If .row < grdWQStd.Rows - 1 Then
                    .row = .row + 1
                Else
                    .row = .row 'again, if the row is on the last row, don't move cells
                End If
                KeyMoveUpdate
            Case vbKeyRight 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                txtActiveCell.Visible = False
                .SetFocus
                .col = .col + 1
        End Select
    End With


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_LostFocus()
  On Error GoTo ErrorHandler

    With grdWQStd
        If intCol = .col And intRow = .row Then
            .Text = txtActiveCell.Text
             txtActiveCell.Visible = False
            .SetFocus
        End If
    End With
    
    If CheckText(strUndoText, txtActiveCell.Text) > 0 Then
        cmdSave.Enabled = True
    End If
    
  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'Counter to determine if changes have been made
Private Function CheckText(strOriginal As String, strNew As String) As Integer
  On Error GoTo ErrorHandler

    
    If strOriginal <> strNew Then
        intNumChanges = intNumChanges + 1
        CheckText = intNumChanges
    End If

  Exit Function
ErrorHandler:
  HandleError False, "CheckText " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub txtWQStdDesc_Click()
  On Error GoTo ErrorHandler

    m_bolChange = True
    cmdSave.Enabled = m_bolChange


  Exit Sub
ErrorHandler:
  HandleError True, "txtWQStdDesc_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub KeyMoveUpdate()

On Error GoTo ErrorHandler

'This guy basically replicates the functionality of the grdWQstd_Click event and is used
'in a couple of instances for moving around the grid.

    m_intWQRow = grdWQStd.row
    m_intWQCol = grdWQStd.col

    txtActiveCell.Visible = False

    If (m_intWQCol = 2) And (m_intWQRow > 0) Then

        With txtActiveCell

            .Move grdWQStd.Left + grdWQStd.CellLeft, _
            grdWQStd.Top + grdWQStd.CellTop, _
            grdWQStd.CellWidth - 30, _
            grdWQStd.CellHeight - 150

            If m_intWQRow <> 0 Then

                .Visible = True
                m_strUndoText = grdWQStd.TextMatrix(m_intWQRow, m_intWQCol)
                .Text = m_strUndoText

                If IsNumeric(m_strUndoText) Then
                    txtActiveCell.Alignment = 1
                Else
                    txtActiveCell.Alignment = 0
                End If

            End If
        End With
        txtActiveCell.SetFocus

    ElseIf m_intWQCol = 0 Then

        txtActiveCell.Visible = False

    End If
    
  Exit Sub
ErrorHandler:
  HandleError False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub UpdateData()
  On Error GoTo ErrorHandler


    Dim strSQLWQStd As String
    Dim strWQSelect As String
    Dim rsWQstd As New ADODB.Recordset
    Set rsWQStdCboClick = New ADODB.Recordset
    Dim rsPollUpdate As New ADODB.Recordset
    Dim i As Integer
    
    Dim booYesNo As Integer
     
    'Selection based on combo box, update Description
    strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"
    
    With rsWQStdCboClick
        .CursorLocation = adUseClient
        .Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    End With
    
    rsWQStdCboClick!Description = txtWQStdDesc.Text
    rsWQStdCboClick.Update
    
    'Now update Threshold values
    For i = 1 To grdWQStd.Rows - 1
        strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.TextMatrix(i, 3)
        
        rsWQstd.Open strWQSelect, g_ADOConn, adOpenDynamic, adLockOptimistic
        
        rsWQstd!Threshold = grdWQStd.TextMatrix(i, 2)
        rsWQstd.Update
        rsWQstd.Close
    Next i

    m_bolChange = False
    
    'Cleanup
    Set rsWQstd = Nothing
    rsWQStdCboClick.Close
    Set rsWQStdCboClick = Nothing
    
    


  Exit Sub
ErrorHandler:
  HandleError False, "UpdateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function ValidateData() As Boolean
  On Error GoTo ErrorHandler

    
    Dim i As Integer
    
    For i = 1 To grdWQStd.Rows - 1
        
        If IsNumeric(grdWQStd.TextMatrix(i, 2)) Then
            If CInt(grdWQStd.TextMatrix(i, 2)) >= 0 Then
                ValidateData = True
            Else
                MsgBox "Warning: Values must be greater than or equal to 0.", vbCritical, "Invalid Value"
                grdWQStd.row = i
                grdWQStd.col = 2
                KeyMoveUpdate
                ValidateData = False
            End If
        ElseIf grdWQStd.TextMatrix(i, 2) <> "" Then
            MsgBox "Numeric values only please.", vbCritical, "Numeric Values Only"
            grdWQStd.row = i
            grdWQStd.col = 2
            KeyMoveUpdate
            ValidateData = False
        End If
    Next
    
    
    


  Exit Function
ErrorHandler:
  HandleError False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

