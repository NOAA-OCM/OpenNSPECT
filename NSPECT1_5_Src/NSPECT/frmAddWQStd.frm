VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAddWQStd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Water Quality Standard"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtActiveCell 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1215
      TabIndex        =   7
      Top             =   3975
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtWQStdDesc 
      Height          =   300
      Left            =   1545
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   4290
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4650
      TabIndex        =   3
      Top             =   3915
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3630
      TabIndex        =   2
      Top             =   3915
      Width           =   975
   End
   Begin VB.TextBox txtWQStdName 
      Height          =   300
      Left            =   1545
      TabIndex        =   0
      Top             =   195
      Width           =   2000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWQStd 
      Height          =   2610
      Left            =   315
      TabIndex        =   6
      Top             =   1140
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   4604
      _Version        =   393216
      Cols            =   3
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483640
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   285
      TabIndex        =   5
      Top             =   615
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Standard Name"
      Height          =   255
      Index           =   7
      Left            =   285
      TabIndex        =   4
      Top             =   255
      Width           =   1245
   End
End
Attribute VB_Name = "frmAddWQStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  frmAddWQStd
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey  -  ed.dempsey@noaa.gov
' *************************************************************************************
' *  Description:  Form that allows for the addition of new water quality standard
' *
' *
' *  Called By:  frmWQStd
' *************************************************************************************

Option Explicit

Private intRow As Integer           'Integer tracking current row
Private intCol As Integer           'Integer tracking current col
Private strUndoText As String       'Text of txtActiveCell
Private mChange As Boolean
Private intNumChanges As Integer

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmAddWQStd.frm"


Private Sub cmdSave_Click()
  On Error GoTo ErrorHandler

    Dim strName As String
    Dim strDescript As String
    Dim strCmd As String
    
    'Get rid of possible apostrophes
    strName = Replace(Trim(txtWQStdName.Text), "'", "''")
    strDescript = Trim(txtWQStdDesc.Text)
    
    If Len(strName) = 0 Then
        MsgBox "Please enter a name for the water quality standard.", vbCritical, "Empty Name Field"
        txtWQStdName.SetFocus
        Exit Sub
    Else
        'Name Check
        If modUtil.UniqueName("WQCRITERIA", txtWQStdName.Text) Then
             'Value check
             If CheckThreshValues Then
                  strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & _
                        Replace(txtWQStdName.Text, "'", "''") & "', '" & _
                        Replace(strDescript, "'", "''") & "')"
                  g_ADOConn.Execute strCmd, adCmdText
             Else
                MsgBox "Threshold values must be numeric.", vbCritical, "Check Threshold Value"
                Exit Sub
             End If
        Else
             MsgBox "The name you have chosen is already in use.  Please select another.", vbCritical, "Select Unique Name"
             Exit Sub
        End If
    End If
    
    'If it gets here, time to add the pollutants
    Dim i As Integer
    i = 0
    
    For i = 1 To grdWQStd.Rows - 1
        'Allow for numeric or nulls
        PollutantAdd txtWQStdName.Text, grdWQStd.TextMatrix(i, 1), grdWQStd.TextMatrix(i, 2)
    Next i
        
    MsgBox txtWQStdName.Text & " successfully added.", vbOKOnly, "Record Added"
    
    'Clean up stuff
    If modUtil.IsLoaded("frmWQStd") Then
        frmWQStd.cboWQStdName.Clear
        modUtil.InitComboBox frmWQStd.cboWQStdName, "WQCRITERIA"
        frmWQStd.cboWQStdName.ListIndex = modUtil.GetCboIndex(txtWQStdName.Text, frmWQStd.cboWQStdName)
    ElseIf modUtil.IsLoaded("frmPrj") Then
        frmPrj.cboWQStd.Clear
        modUtil.InitComboBox frmPrj.cboWQStd, "WQCRITERIA"
        frmPrj.cboWQStd.AddItem "Define a new water quality standard...", frmPrj.cboWQStd.ListCount
        frmPrj.cboWQStd.ListIndex = modUtil.GetCboIndex(txtWQStdName.Text, frmPrj.cboWQStd)
    End If
    
    Unload Me
    
  Exit Sub
ErrorHandler:
  HandleError True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function CheckThreshValues() As Boolean
  On Error GoTo ErrorHandler

    Dim i As Integer
    i = 0
    
    For i = 1 To grdWQStd.Rows - 1
        If IsNumeric(grdWQStd.TextMatrix(i, 2)) Or grdWQStd.TextMatrix(i, 2) = "" Then
            CheckThreshValues = True
        Else
            CheckThreshValues = False
            Exit Function
        End If
    Next i
    
  Exit Function
ErrorHandler:
  HandleError False, "CheckThreshValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    
    'Populate cbo with pollutant names
    Dim rsPollutant As ADODB.Recordset
    Dim strPollutant As String
    Set rsPollutant = New ADODB.Recordset
    
    mChange = False
    intNumChanges = 0
    
    modUtil.InitWQStdGrid grdWQStd
    
    strPollutant = "SELECT NAME FROM POLLUTANT ORDER BY NAME ASC"
    
    rsPollutant.CursorLocation = adUseClient
    rsPollutant.Open strPollutant, g_ADOConn, adOpenDynamic, adLockOptimistic
     
    Dim i As Integer
    i = 1
     
    rsPollutant.MoveFirst
    grdWQStd.Rows = rsPollutant.RecordCount + 1
       
    For i = 1 To rsPollutant.RecordCount
        grdWQStd.TextMatrix(i, 1) = rsPollutant!Name
        rsPollutant.MoveNext
    Next i
    
    'Cleanup
    rsPollutant.Close
    Set rsPollutant = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", vbYesNo, "Exit?")
    
    If intvbYesNo = vbYes Then
        If IsLoaded("frmPrj") Then
            frmPrj.cboWQStd.ListIndex = 0
        End If
        
        Unload frmAddWQStd
    Else
        Exit Sub
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
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

Private Sub txtActiveCell_Change()
  On Error GoTo ErrorHandler

    'See the grd text to the text box
    grdWQStd.Text = txtActiveCell.Text


  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_GotFocus()
'Select all when clicked
  On Error GoTo ErrorHandler

    strUndoText = txtActiveCell.Text
    
    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)

  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Key event handles for txtActiveCell
Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler

    With grdWQStd
    
    Select Case KeyCode
        Case vbKeyEscape 'if the user pressed escape, then get out without changing
            .Text = strUndoText
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

Private Sub KeyMoveUpdate()
  On Error GoTo ErrorHandler

'This guy basically replicates the functionality of the grdWQstd_Click event and is used
'in a couple of instances for moving around the grid.

    intRow = grdWQStd.row
    intCol = grdWQStd.col

    txtActiveCell.Visible = False

    If (intCol = 2) And (intRow >= 1) Then

        With txtActiveCell

            .Move grdWQStd.Left + grdWQStd.CellLeft, _
            grdWQStd.Top + grdWQStd.CellTop, _
            grdWQStd.CellWidth - 30, _
            grdWQStd.CellHeight - 150

            If intRow <> 0 Then

                .Visible = True
                strUndoText = grdWQStd.TextMatrix(intRow, intCol)
                .Text = strUndoText

            End If
        End With
        txtActiveCell.SetFocus

    ElseIf intCol = 0 Then

        txtActiveCell.Visible = False

    End If



  Exit Sub
ErrorHandler:
  HandleError False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
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
    
    If CheckText(strUndoText, grdWQStd.Text) > 0 Then
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


Private Sub PollutantAdd(strName As String, strPoll As String, intThresh As String)
  On Error GoTo ErrorHandler

    Dim strPollAdd As String
    Dim strPollDetails As String
    Dim strCmdInsert As String
    
    Dim rsPollAdd As ADODB.Recordset
    Dim rsPollDetails As ADODB.Recordset
    
    Set rsPollAdd = New ADODB.Recordset
    Set rsPollDetails = New ADODB.Recordset
       
    'Get the WQCriteria values using the name
    strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
    rsPollAdd.CursorLocation = adUseClient
    rsPollAdd.Open strPollAdd, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'Get the pollutant particulars
    strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
    rsPollDetails.CursorLocation = adUseClient
    rsPollDetails.Open strPollDetails, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    If Trim(intThresh) = "" Then
        strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID) VALUES ('" & _
                rsPollDetails!POLLID & "', '" & _
                rsPollAdd!WQCRITID & "')"
    Else
        strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & _
                rsPollDetails!POLLID & "', '" & _
                rsPollAdd!WQCRITID & "'," & _
                intThresh & ")"
    End If
    
    g_ADOConn.Execute strCmdInsert, adCmdText
       
    Set rsPollAdd = Nothing
    Set rsPollDetails = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtWQStdName_Change()
  On Error GoTo ErrorHandler

    txtWQStdName.Text = Replace(txtWQStdName.Text, "'", "")

  Exit Sub
ErrorHandler:
  HandleError True, "txtWQStdName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
