VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNewLCType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Land Cover Type"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWWL 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4590
      TabIndex        =   11
      Top             =   5085
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtActiveCell 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2295
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4980
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLcClasses 
      Height          =   3465
      Left            =   180
      TabIndex        =   9
      ToolTipText     =   "Right click to add, delete, or insert a row"
      Top             =   1395
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   6112
      _Version        =   393216
      Cols            =   11
      ForeColorSel    =   -2147483640
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.TextBox txtLCTypeDesc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   630
      Width           =   5600
   End
   Begin VB.TextBox txtLCType 
      Height          =   285
      Left            =   1815
      TabIndex        =   0
      Top             =   225
      Width           =   2000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6510
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7635
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Index           =   6
      Left            =   375
      TabIndex        =   8
      Top             =   630
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SCS Curve Numbers"
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      Top             =   1080
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Classification"
      Height          =   285
      Index           =   1
      Left            =   165
      TabIndex        =   6
      Top             =   1080
      Width           =   3405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUSLE"
      Height          =   285
      Index           =   2
      Left            =   6810
      TabIndex        =   5
      Top             =   1080
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Land Cover Type:"
      Height          =   255
      Index           =   7
      Left            =   195
      TabIndex        =   4
      Top             =   255
      Width           =   1410
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAppendRow 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuInsertRow 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "frmNewLCType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmNewLCType
' *************************************************************************************
' *  Description: Form for entering a new Land Class Type
' *  within NSPECT
' *
' *  Called By:  frmLCType menu New...
' *************************************************************************************

Option Explicit

Private m_intRow As Integer                   'Current Row
Private m_intCol As Integer                   'Current Col.
Private m_intLCTypeID As Long                 'LCTypeID#

Private m_bolGridChanged As Boolean           'Flag for whether or not grid values have changed
Private m_bolSaved As Boolean                 'Flag for saved/not saved changes
Private m_bolFirstLoad As Boolean             'Is initial Load event
Private m_bolBegin As Boolean
Private m_intCount As Integer

Private m_strUndoText As String               'initial cell value - defaults back on Esc

Private m_intMouseButton As Integer           'Integer for mouse button click - added to avoid right click change cell value problem
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmNewLCType.frm"


Private Sub chkWWL_Click(Index As Integer)
  On Error GoTo ErrorHandler


    grdLCClasses.TextMatrix(Index, 8) = chkWWL(Index).Value



  Exit Sub
ErrorHandler:
  HandleError True, "chkWWL_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

    Dim strName As String
    Dim strDescript As String
    Dim strCmd As String            'INSERT function
    Dim arrParams(7) As Variant     'Array that holds each row's contents
    Dim i As Integer
    Dim j As Integer
    
    
    If ValidateGridValues Then
        'Get rid of possible apostrophes in name
        strName = Trim(txtLCType.Text)
        strDescript = Trim(txtLCTypeDesc.Text)
        
        If modUtil.UniqueName("LCTYPE", strName) And (Trim(strName) <> "") Then
            strCmd = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & _
                    Replace(strName, "'", "''") & "', '" & _
                    Replace(strDescript, "'", "''") & "')"
            g_ADOConn.Execute strCmd, adCmdText
        Else
            MsgBox "The name you have chosen is already in use.  Please select another.", vbCritical, "Select Unique Name"
            Exit Sub
        End If 'End unique name check
        
        'Now add GRID values
        
        For i = 1 To grdLCClasses.Rows - 1
            For j = 1 To 8              'Hard coded, bad, but ya know
                arrParams(j - 1) = grdLCClasses.TextMatrix(i, j)
            Next j
            
            AddLCClass strName, arrParams
            
        Next i
        
        MsgBox "Data saved successfully.", vbOKOnly, "Data Saved"
        
        Unload Me
        Unload frmLCTypes
    Else
        Exit Sub
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub AddLCClass(strName As String, strParams())
    'Called from cmdOK_Click, this uses a passed array to insert new landclasses
    
    Dim strLCTypeAdd As String
    Dim strCmdInsert As String
    
    Dim rsLCType As ADODB.Recordset
    Set rsLCType = New ADODB.Recordset
    
On Error GoTo ErrHandler:
    'Get the WQCriteria values using the name
    strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
    rsLCType.CursorLocation = adUseClient
    rsLCType.Open strLCTypeAdd, g_ADOConn, adOpenDynamic, adLockOptimistic
               
    'On Error GoTo dbaseerr
    strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & _
            CStr(strParams(0)) & ",'" & _
            CStr(strParams(1)) & "'," & _
            CStr(rsLCType!LCTypeID) & "," & _
            CStr(strParams(2)) & "," & _
            CStr(strParams(3)) & "," & _
            CStr(strParams(4)) & "," & _
            CStr(strParams(5)) & "," & _
            CStr(strParams(6)) & "," & _
            CStr(strParams(7)) & ")"
  
    g_ADOConn.Execute strCmdInsert, adCmdText
    
    Set rsLCType = Nothing

    Exit Sub
    
ErrHandler:
    MsgBox "There is an error inserting records into LCClass.", vbCritical, Err.Number & ": " & Err.Description
    Exit Sub
    
End Sub


Private Sub Form_GotFocus()
  On Error GoTo ErrorHandler
    
    txtLCType.SetFocus

  Exit Sub
ErrorHandler:
  HandleError True, "Form_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

   
    InitLCClassesGrid grdLCClasses
    
    With grdLCClasses
       .row = 1
       .TextMatrix(.row, 1) = "0"
       .TextMatrix(.row, 2) = "Landclass" & .row
       .TextMatrix(.row, 3) = "0"
       .TextMatrix(.row, 4) = "0"
       .TextMatrix(.row, 5) = "0"
       .TextMatrix(.row, 6) = "0"
       .TextMatrix(.row, 7) = "0"
       .TextMatrix(.row, 8) = "0"
    End With
    
    CreateCheckBoxes True
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler
  
   Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Click()
  On Error GoTo ErrorHandler

    txtActiveCell.Visible = False

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler
  
    Unload Me

  Exit Sub
ErrorHandler:
  HandleError False, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub grdLCClasses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

    m_intRow = grdLCClasses.MouseRow
    m_intCol = grdLCClasses.MouseCol

    m_intMouseButton = Button
    
    If Button = 2 Then
        
        If m_intCol = 0 Then
            With grdLCClasses
                .row = m_intRow
                .col = m_intCol
                modUtil.HighlightGridRow grdLCClasses, m_intRow
            End With
            PopupMenu mnuPopUp
        End If
        
    End If

  Exit Sub
ErrorHandler:
  HandleError False, "grdLCClasses_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub grdLcClasses_Click()
  On Error GoTo ErrorHandler

    m_intRow = grdLCClasses.MouseRow
    m_intCol = grdLCClasses.MouseCol

    txtActiveCell.Visible = False
    
    If m_intMouseButton = 1 Then 'Right clicking snagged cell values, so we can only act on left click event
    
        If (m_intCol >= 1 And m_intCol <= 8) And (m_intRow >= 1) Then
            
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
                    
                    If IsNumeric(m_strUndoText) Then
                        txtActiveCell.Alignment = 1
                    Else
                        txtActiveCell.Alignment = 0
                    End If
                End If
            End With
        
        ElseIf m_intCol = 0 Then
            
            txtActiveCell.Visible = False
    
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "grdLcClasses_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub txtActiveCell_Change()
  On Error GoTo ErrorHandler

    grdLCClasses.Text = txtActiveCell.Text
    m_bolGridChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtActiveCell_GotFocus()
  On Error GoTo ErrorHandler
    
    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)

  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler

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
                    .col = .col
                    .Text = txtActiveCell.Text
                    txtActiveCell.Visible = False
                End If
            Case vbKeyRight 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                '.Text = txtActiveCell.Text
                txtActiveCell.Visible = False
                If .col < 8 Then
                    .col = .col + 1
                    KeyMoveUpdate
                Else
                    .col = .col
                    .Text = txtActiveCell.Text
                    txtActiveCell.Visible = False
                End If
            Case vbKeyTab
                MsgBox "WWWWWWWWWWWOOOOOO"
        End Select
    End With

  Exit Sub
ErrorHandler:
  HandleError True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub KeyMoveUpdate()
  On Error GoTo ErrorHandler

'This guy basically replicates the functionality of the grdLCClasses_Click event and is used
'in a couple of instances for moving around the grid.
 
    m_intRow = grdLCClasses.row
    m_intCol = grdLCClasses.col

    'cboWWL.Visible = False
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
                
                If IsNumeric(m_strUndoText) Then
                    txtActiveCell.Alignment = 1
                Else
                    txtActiveCell.Alignment = 0
                End If
                
            End If
        End With
        txtActiveCell.SetFocus
    
    ElseIf m_intCol = 8 And m_intRow <> 0 Then
          
    ElseIf m_intCol = 0 Then
        
        txtActiveCell.Visible = False
    
    End If
    
  Exit Sub
ErrorHandler:
  HandleError False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuAppendRow_Click()
  On Error GoTo ErrorHandler

    'add row to end of grid
    With grdLCClasses
       .Rows = .Rows + 1
       .row = .Rows - 1
       .TextMatrix(.row, 1) = "0"
       .TextMatrix(.row, 2) = "Landclass" & .row
       .TextMatrix(.row, 3) = "0"
       .TextMatrix(.row, 4) = "0"
       .TextMatrix(.row, 5) = "0"
       .TextMatrix(.row, 6) = "0"
       .TextMatrix(.row, 7) = "0"
       .TextMatrix(.row, 8) = "0"
    End With
    
    CreateCheckBoxes False, grdLCClasses.row

    m_bolGridChanged = True       'reset
    cmdOK.Enabled = m_bolGridChanged
   
  Exit Sub
ErrorHandler:
  HandleError True, "mnuAppendRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuDeleteRow_Click()
  On Error GoTo ErrorHandler

'delete current row
    Dim R%, C%, row%, col%

    With grdLCClasses
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
              .Rows = .Rows - 1     'del last row
          End If
      End If
   End With
    
   m_intCount = grdLCClasses.Rows
   ClearCheckBoxes True, m_intCount + 1
   CreateCheckBoxes True
    
  Exit Sub
ErrorHandler:
  HandleError True, "mnuDeleteRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuInsertRow_Click()
  On Error GoTo ErrorHandler

    'insert row above cuurent row
    Dim R%, row%, col%
    With grdLCClasses
      If .row < .FixedRows Then        'make sure we don't insert above header Rows
         mnuAppendRow_Click
      Else
         R% = .row
         .Rows = .Rows + 1               'add a row
         '.TextMatrix(.Rows - 1, 0) = .Rows - .FixedRows     'new row Title

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
     End If
   End With

   txtActiveCell.Visible = False

   m_intCount = grdLCClasses.Rows
   
   ClearCheckBoxes True, m_intCount - 1
   CreateCheckBoxes True

  Exit Sub
ErrorHandler:
  HandleError True, "mnuInsertRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function ValidateGridValues() As Boolean
  On Error GoTo ErrorHandler

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
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
                
            Case 4
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
                
            Case 5
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
            
            Case 6
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
                    ErrorGenerator Err1, i, j
                    grdLCClasses.col = j
                    grdLCClasses.row = i
                    ValidateGridValues = False
                    KeyMoveUpdate
                    Exit Function
                End If
            
            Case 7
                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Then
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

  Exit Function
ErrorHandler:
  HandleError False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub CreateCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler

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
            Set chkWWL(k).Container = frmNewLCType
            With chkWWL(k)
                .Visible = True
                .Top = grdLCClasses.Top + grdLCClasses.CellTop
                .Left = grdLCClasses.Left + grdLCClasses.CellLeft + (grdLCClasses.CellWidth * 0.4)
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
            Set chkWWL(k).Container = frmNewLCType
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

  Exit Sub
ErrorHandler:
  HandleError False, "CreateCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ClearCheckBoxes(booAll As Boolean, intCount As Integer, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler

    Dim k As Integer
    
    If booAll Then
    
        For k = 1 To intCount - 1
            Unload chkWWL(k)
        Next k
    Else
        Unload chkWWL(intRecNo)
    End If

  Exit Sub
ErrorHandler:
  HandleError False, "ClearCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

