VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLUScen 
   Caption         =   "Edit Land Use Scenario"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSelectedPolys 
      Caption         =   "Use Selected Polygons Only"
      Height          =   240
      Left            =   2055
      TabIndex        =   24
      Top             =   930
      Width           =   3465
   End
   Begin VB.TextBox txtTypes 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2265
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox cboPollName 
      Height          =   315
      Left            =   255
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   4905
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CheckBox chkWatWetlands 
      Caption         =   "Water/Wetlands"
      Height          =   285
      Left            =   3465
      TabIndex        =   7
      Top             =   2055
      Width           =   1635
   End
   Begin VB.TextBox txtLUName 
      Height          =   285
      Left            =   2055
      MaxLength       =   30
      TabIndex        =   0
      Top             =   105
      Width           =   3510
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4935
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4935
      Width           =   975
   End
   Begin VB.TextBox txtLUCN 
      Height          =   285
      Index           =   4
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   6
      Text            =   "0"
      Top             =   2070
      Width           =   900
   End
   Begin VB.TextBox txtLUCN 
      Height          =   285
      Index           =   3
      Left            =   4740
      MaxLength       =   30
      TabIndex        =   5
      Text            =   "0"
      Top             =   1545
      Width           =   900
   End
   Begin VB.TextBox txtLUCN 
      Height          =   285
      Index           =   2
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   4
      Text            =   "0"
      Top             =   1545
      Width           =   900
   End
   Begin VB.TextBox txtLUCN 
      Height          =   285
      Index           =   1
      Left            =   2955
      MaxLength       =   30
      TabIndex        =   3
      Text            =   "0"
      Top             =   1545
      Width           =   900
   End
   Begin VB.TextBox txtLUCN 
      Height          =   285
      Index           =   0
      Left            =   2055
      MaxLength       =   30
      TabIndex        =   2
      Text            =   "0"
      Top             =   1545
      Width           =   900
   End
   Begin VB.ComboBox cboLULayer 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   525
      Width           =   3555
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPoll 
      Height          =   1860
      Left            =   165
      TabIndex        =   23
      Top             =   2880
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   3281
      _Version        =   393216
      Cols            =   6
      BackColorSel    =   12648447
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      Height          =   255
      Index           =   5
      Left            =   4740
      TabIndex        =   20
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      Height          =   255
      Index           =   3
      Left            =   3855
      TabIndex        =   19
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      Height          =   255
      Index           =   2
      Left            =   2955
      TabIndex        =   18
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Scenario Name:"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   17
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Cover Factor:"
      Height          =   285
      Index           =   15
      Left            =   360
      TabIndex        =   16
      Top             =   2100
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      Height          =   255
      Index           =   4
      Left            =   2055
      TabIndex        =   15
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "SCS Curve Numbers:"
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   14
      Top             =   1305
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   16
      Left            =   570
      TabIndex        =   13
      Top             =   2550
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coefficients"
      Height          =   255
      Index           =   17
      Left            =   2565
      TabIndex        =   12
      Top             =   2550
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   18
      Left            =   165
      TabIndex        =   11
      Top             =   2550
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Layer:"
      Height          =   285
      Index           =   19
      Left            =   315
      TabIndex        =   10
      Top             =   510
      Width           =   495
   End
   Begin VB.Menu mnuPopLU 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuAppendP 
         Caption         =   "Append Row"
      End
      Begin VB.Menu mnuInsertP 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDeleteP 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "frmLUScen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmLUScen
' *************************************************************************************
' *  Description: Form that takes in the user's stuff for a new landuse scenario.  Stores
' *  the contents in an xml string in a hidden cell in 3rd column in grdLU on frmPrj.
' *  This is so folks can make use of the Edit Scenario... menu choice as well.
' *
' *  Called By: frmPrj::Add/Edit Landuse scenario
' *************************************************************************************
' *  Subs:
' *     Function CreateXMLFile:: creates XML string holding the forms info
' *     Function ValidateData:: You guessed it.  Validates data.
' *     Sub PopulateForm:: Throws values of the hidden cell in grdLU into form
' *
' *  Misc: employs clsPopUp, an API use of a popup menu.  Workaround for VB does not
' *  support multiple popups on the same form.
' *************************************************************************************

Option Explicit

Private m_App As IApplication
Private m_pMap As IMap
Private m_intRow As Integer
Private m_intCol As Integer
Private m_strUndoText As String
Private clsManScen As clsXMLLUScen
Private m_strWQStd As String

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmLUScen.frm"

Private Sub cboPollName_Click()

If m_intRow > 0 Then
    grdPoll.TextMatrix(m_intRow, m_intCol) = cboPollName.Text
    cboPollName.Visible = False
End If

End Sub

Private Sub cmdCancel_Click()
   
   Unload frmLUScen

End Sub

Private Sub cmdOK_Click()

    If ValidateData Then   'ToDO: Validate data
        clsManScen.SaveFile CreateXMLFile
        Unload frmLUScen
    End If
    
End Sub

Private Function CreateXMLFile() As String

   Dim i As Integer
   Dim clsMan As New clsXMLLUScen
   
   With clsMan
    
        .strLUScenName = Trim(txtLUName.Text)
        .strLUScenLyrName = Trim(cboLULayer.Text)
        .strLUScenFileName = modUtil.GetFeatureFileName(.strLUScenLyrName, m_App)
        .intLUScenSelectedPoly = chkSelectedPolys.Value
        .intSCSCurveA = txtLUCN(0).Text
        .intSCSCurveB = txtLUCN(1).Text
        .intSCSCurveC = txtLUCN(2).Text
        .intSCSCurveD = txtLUCN(3).Text
        .lngCoverFactor = txtLUCN(4).Text
        .intWaterWetlands = chkWatWetlands.Value
        
        For i = 1 To grdPoll.Rows - 1
        Set clsMan.clsPollutant = New clsXMLLUScenPollItem
            .clsPollutant.intID = i
            .clsPollutant.strPollName = grdPoll.TextMatrix(i, 1)
            .clsPollutant.intType1 = CDbl(grdPoll.TextMatrix(i, 2))
            .clsPollutant.intType2 = CDbl(grdPoll.TextMatrix(i, 3))
            .clsPollutant.intType3 = CDbl(grdPoll.TextMatrix(i, 4))
            .clsPollutant.intType4 = CDbl(grdPoll.TextMatrix(i, 5))
           .clsPollItems.Add .clsPollutant
        Next i
        
    End With
    
    frmPrj.grdLU.TextMatrix(g_intManScenRow, 2) = clsMan.strLUScenName
    CreateXMLFile = clsMan.XML

End Function


Private Sub Form_Load()
   
   modUtil.AddFeatureLayerToComboBox cboLULayer, m_pMap, "poly"
   FillGrid
   
  'define flexgrid for Land Use tab
   With grdPoll
      .col = .FixedCols
      .row = .FixedRows
      .Width = 5600 + (.GridLineWidth * (.Cols + 1)) + 75
      .ColWidth(0) = 400
      .ColWidth(1) = 2000
      .ColWidth(2) = 800
      .ColWidth(3) = 800
      .ColWidth(4) = 800
      .ColWidth(5) = 800
      .row = 0
      .col = 0
      .Text = ""
      .col = 1
      .Text = "Pollutant"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Type 1"
      .CellAlignment = flexAlignCenterCenter
      .col = 3
      .Text = "Type 2"
      .CellAlignment = flexAlignCenterCenter
      .col = 4
      .Text = "Type 3"
      .CellAlignment = flexAlignCenterCenter
      .col = 5
      .Text = "Type 4"
      .CellAlignment = flexAlignCenterCenter
      .Enabled = True
   End With
   
   Set clsManScen = New clsXMLLUScen
   
   If Len(g_strLUScenFileName) > 0 Then
        clsManScen.XML = g_strLUScenFileName
        grdPoll.Rows = clsManScen.clsPollItems.Count + 1
        PopulateForm
   Else
        txtLUCN(0).Text = "0"
        txtLUCN(1).Text = "0"
        txtLUCN(2).Text = "0"
        txtLUCN(3).Text = "0"
   End If

End Sub

Public Sub init(ByVal pApp As IApplication, strWQStd As String)
  
    Dim pMxDoc As IMxDocument
    Set m_App = pApp
    
    Set pMxDoc = m_App.Document
    Set m_pMap = pMxDoc.FocusMap
    
    m_strWQStd = strWQStd
    

 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set clsManScen = Nothing
    m_intRow = 0
    m_intCol = 0
    
End Sub

Private Sub grdPoll_Click()

    m_intRow = grdPoll.MouseRow
    m_intCol = grdPoll.MouseCol
    
    If m_intCol > 1 And m_intRow >= 1 Then
        With txtTypes
            .Visible = True
            .Move grdPoll.Left + grdPoll.CellLeft, _
            grdPoll.Top + grdPoll.CellTop
            .Width = grdPoll.CellWidth - 30
            .Height = grdPoll.CellHeight - 30
            .Text = grdPoll.TextMatrix(m_intRow, m_intCol)
        End With
        
    Else
        If m_intCol < 1 Then
            txtTypes.Visible = False
            cboPollName.Visible = False
            
        End If
    End If
    
End Sub

Private Sub grdPoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
    m_intRow = grdPoll.MouseRow
    m_intCol = grdPoll.MouseCol

  Exit Sub
ErrorHandler:
  HandleError True, "grdPoll_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub AlterGrid(lngAlterType As Long)
    Dim R%, C%, row%, col%
    
    Select Case lngAlterType
        Case 1  'Append
            
            With grdPoll
               .Rows = .Rows + 1
               .row = .Rows - 1
            End With
            
        Case 2  'Insert
            
            With grdPoll
                If .row < .FixedRows Then        'make sure we don't insert above header Rows
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                Else
                   R% = .row
                   .Rows = .Rows + 1             'add a row
            
                   For row% = .Rows - 1 To R% + 1 Step -1 'move data dn 1 row
                       For col% = 1 To .Cols - 1
                           .TextMatrix(row%, col%) = .TextMatrix(row% - 1, col%)
                       Next col%
                   Next row%
                   For col% = 1 To .Cols - 1       ' clear all cells in this row
                      .TextMatrix(R%, col%) = ""
                   Next col%
                End If
            End With
         
         Case 3     'Delete
            
            Dim lngLCClassID As Long
    
            With grdPoll
             
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
                   End If
               End If
            End With
       End Select

End Sub

Private Sub txtTypes_Change()
    
    grdPoll.TextMatrix(m_intRow, m_intCol) = txtTypes.Text
    
End Sub

Private Sub txtTypes_GotFocus()
    
    txtTypes.SelStart = 0
    txtTypes.SelLength = Len(txtTypes.Text)
    m_strUndoText = txtTypes.Text
    
End Sub

Private Sub txtTypes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With grdPoll
        Select Case KeyCode
            Case vbKeyEscape 'if the user pressed escape, then get out without changing
                .Text = m_strUndoText
                txtTypes.Visible = False
                .SetFocus
                
            Case 13 'if the user presses enter, get out of the textbox
                .Text = txtTypes.Text
                txtTypes.Visible = False
            
            Case vbKeyUp 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
                .Text = txtTypes.Text
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
                .Text = txtTypes.Text
                txtTypes.Visible = False
                If .row < grdPoll.Rows - 1 Then
                    .row = .row + 1
                    KeyMoveUpdate
                Else
                    If .col > 1 Then
                        .row = 1 'again, if the row is on the last row, don't move cells
                        .col = .col - 1
                    Else
                        .Text = txtTypes.Text
                        txtTypes.Visible = False
                    End If
                    KeyMoveUpdate
                End If
            Case vbKeyLeft 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                txtTypes.Visible = False
                If .col >= 2 Then
                    .col = .col - 1
                    KeyMoveUpdate
                    
                Else
                
                If .row > 1 And .col = 2 Then
                    .col = 5
                    .row = .row - 1
                    KeyMoveUpdate
                End If
                End If
                
            Case vbKeyRight 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
                .Text = txtTypes.Text
                txtTypes.Visible = False
                If .col < 5 Then
                    .col = .col + 1
                    KeyMoveUpdate
                Else
                    If .row < grdPoll.Rows - 1 Then
                        .col = 2
                        .row = .row + 1
                        KeyMoveUpdate
                    End If
                End If
        End Select
    End With
    
    
End Sub

Private Sub KeyMoveUpdate()

'This guy basically replicates the functionality of the grdpolldef_Click event and is used
'in a couple of instances for moving around the grid.
 
    m_intRow = grdPoll.row
    m_intCol = grdPoll.col

    txtTypes.Visible = False

    If (m_intCol >= 2 And m_intCol <= 5) And (m_intRow >= 1) Then
        
        With txtTypes
        
             .Move grdPoll.Left + grdPoll.CellLeft, _
            grdPoll.Top + grdPoll.CellTop, _
            grdPoll.CellWidth - 30, _
            grdPoll.CellHeight - 150
             
            If m_intRow <> 0 Then
                
                .Visible = True
                 m_strUndoText = grdPoll.TextMatrix(m_intRow, m_intCol)
                .Text = m_strUndoText
                .Alignment = 1
            End If
        End With
        
        txtTypes.SetFocus
    
    ElseIf m_intCol = 0 Then
        
        txtTypes.Visible = False
    
    End If
    
End Sub

Private Function ValidateData() As Boolean

    Dim i As Integer
    
    'Project Name
    If Len(txtLUName.Text) = 0 Or Len(txtLUName.Text) > 30 Then
        MsgBox "Please enter a name for the scenario.  Names must be 30 characters or less.", vbCritical, "Enter Name"
        txtLUName.SetFocus
        ValidateData = False
        Exit Function
    Else
        ValidateData = True
    End If
    
    'LandCoverLayer
    If cboLULayer.Text = "" Then
        MsgBox "Please select a layer before continuing.", vbCritical, "Select Layer"
        cboLULayer.SetFocus
        ValidateData = False
        Exit Function
    Else
        If Not modUtil.LayerInMap(cboLULayer, m_pMap) Then
            MsgBox "The layer you have choosen is not in the current map frame.", vbCritical, "Layer Not Found"
            ValidateData = False
            Exit Function
        End If
    End If
    
    'Check selected polygons
    If chkSelectedPolys.Value = 1 Then
        If m_pMap.SelectionCount = 0 Then
            MsgBox "You have chosen to use selected polygons from " & cboLULayer.Text & ", but the current map contains no selected features." & vbNewLine & _
            "Please select features or N-SPECT will use the entire extent of " & cboLULayer.Text & " to apply this landuse scenario.", vbInformation, "No Selected Features Found"
            ValidateData = False
        End If
    End If
    
    'SCS Curve Numbers
    For i = 0 To txtLUCN().UBound
        If IsNumeric(Trim(txtLUCN(i).Text)) Then
            If CInt(txtLUCN(i).Text) > 0 Or CInt(txtLUCN(i).Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox "SCS Values are to be numeric only in the range of 0 - 1.", vbCritical, "Check SCS Values"
            ValidateData = False
            txtLUCN(i).SetFocus
            Exit Function
        End If
    Next i

End Function

Private Sub PopulateForm()
    
    Dim strScenName As String
    Dim strLyrName As String
    Dim i As Integer
     
    Dim clsPollItem As clsXMLLUScenPollItem
     
    strScenName = clsManScen.strLUScenName
    strLyrName = clsManScen.strLUScenLyrName
    
    txtLUName.Text = strScenName
    
    If modUtil.LayerInMap(strLyrName, m_pMap) Then
        cboLULayer.ListIndex = modUtil.GetCboIndex(strLyrName, cboLULayer)
    End If
    
    chkSelectedPolys.Value = clsManScen.intLUScenSelectedPoly
    
    txtLUCN(0).Text = CStr(clsManScen.intSCSCurveA)
    txtLUCN(1).Text = CStr(clsManScen.intSCSCurveB)
    txtLUCN(2).Text = CStr(clsManScen.intSCSCurveC)
    txtLUCN(3).Text = CStr(clsManScen.intSCSCurveD)
    txtLUCN(4).Text = CStr(clsManScen.lngCoverFactor)
    chkWatWetlands.Value = clsManScen.intWaterWetlands
    
    grdPoll.Rows = clsManScen.clsPollItems.Count + 1
    
    For i = 1 To clsManScen.clsPollItems.Count
        With grdPoll
            .row = clsManScen.clsPollItems.Item(i).intID
            .TextMatrix(.row, 1) = clsManScen.clsPollItems.Item(i).strPollName
            .TextMatrix(.row, 2) = CStr(clsManScen.clsPollItems.Item(i).intType1)
            .TextMatrix(.row, 3) = CStr(clsManScen.clsPollItems.Item(i).intType2)
            .TextMatrix(.row, 4) = CStr(clsManScen.clsPollItems.Item(i).intType3)
            .TextMatrix(.row, 5) = CStr(clsManScen.clsPollItems.Item(i).intType4)
        End With
    Next i

End Sub

'Loads the variety of comboboxes throught the project using combobox and name of table
Private Sub FillGrid()

    Dim strSQLWQStd As String
    Dim rsWQStdCboClick As New ADODB.Recordset
    
    Dim strSQLWQStdPoll As String
    Dim rsWQStdPoll As New ADODB.Recordset
    
    Dim i As Integer
    
    'Selection based on combo box
    strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & m_strWQStd & "'"
    rsWQStdCboClick.CursorLocation = adUseClient
    rsWQStdCboClick.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    If rsWQStdCboClick.RecordCount > 0 Then
        
        strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & _
        "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & _
        "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick!WQCRITID
             
        rsWQStdPoll.CursorLocation = adUseClient
        rsWQStdPoll.Open strSQLWQStdPoll, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
        grdPoll.Rows = rsWQStdPoll.RecordCount + 1
        
        For i = 1 To rsWQStdPoll.RecordCount
            
            grdPoll.TextMatrix(i, 1) = rsWQStdPoll!Name
            grdPoll.TextMatrix(i, 2) = 0
            grdPoll.TextMatrix(i, 3) = 0
            grdPoll.TextMatrix(i, 4) = 0
            grdPoll.TextMatrix(i, 5) = 0
            rsWQStdPoll.MoveNext
            
        Next i
               
        'Clean it
        Set rsWQStdPoll = Nothing
        Set rsWQStdCboClick = Nothing
        
    Else
        
        MsgBox "Warning: There are no water quality standards remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        Set rsWQStdCboClick = Nothing
    End If
    
End Sub
