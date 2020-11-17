VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewPollutants 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Pollutant"
   ClientHeight    =   8490
   ClientLeft      =   4020
   ClientTop       =   1905
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPollutant 
      Height          =   300
      Left            =   1590
      TabIndex        =   0
      Top             =   30
      Width           =   2010
   End
   Begin VB.TextBox txtActiveCellWQStd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   375
      TabIndex        =   14
      Top             =   8115
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtActiveCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   390
      TabIndex        =   13
      Top             =   7785
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7230
      Left            =   240
      TabIndex        =   8
      Top             =   435
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12753
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   7056
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Coefficients"
      TabPicture(0)   =   "frmNewPollutants.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(4)=   "Label1(7)"
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(6)=   "grdPollDef"
      Tab(0).Control(7)=   "cboLCType"
      Tab(0).Control(8)=   "txtCoeffSetDesc"
      Tab(0).Control(9)=   "txtCoeffSet"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Water Quality Standards"
      TabPicture(1)   =   "frmNewPollutants.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdWQStd"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtCoeffSet 
         Height          =   285
         Left            =   -73500
         TabIndex        =   1
         Top             =   525
         Width           =   2000
      End
      Begin VB.TextBox txtCoeffSetDesc 
         Height          =   285
         Left            =   -73500
         TabIndex        =   3
         Top             =   930
         Width           =   5985
      End
      Begin VB.ComboBox cboLCType 
         Height          =   315
         ItemData        =   "frmNewPollutants.frx":0038
         Left            =   -69750
         List            =   "frmNewPollutants.frx":003A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWQStd 
         Height          =   4275
         Left            =   345
         TabIndex        =   15
         Top             =   630
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   5
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPollDef 
         Height          =   5340
         Left            =   -74805
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1635
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   9419
         _Version        =   393216
         Cols            =   7
         BackColorSel    =   12648447
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label Label2 
         Caption         =   "Threshold Units: ug/L"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Coefficient Set:"
         Height          =   255
         Index           =   5
         Left            =   -74730
         TabIndex        =   17
         Top             =   555
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Land Cover Type:"
         Height          =   255
         Index           =   7
         Left            =   -71145
         TabIndex        =   16
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   -74775
         TabIndex        =   12
         Top             =   1320
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coefficients (mg/L)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   -70830
         TabIndex        =   11
         Top             =   1320
         Width           =   3390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -74325
         TabIndex        =   10
         Top             =   1320
         Width           =   3450
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Index           =   6
         Left            =   -74730
         TabIndex        =   9
         Top             =   945
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5670
      TabIndex        =   6
      Top             =   7935
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   7935
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   135
      Top             =   6030
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Pollutant Name:"
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   90
      Width           =   1290
   End
   Begin VB.Menu mnuCoeff 
      Caption         =   "&Coefficients"
      Begin VB.Menu mnuCoeffNewSet 
         Caption         =   "&Add Coefficient Set..."
      End
      Begin VB.Menu mnuCoeffCopySet 
         Caption         =   "&Copy Coefficient Set..."
      End
   End
End
Attribute VB_Name = "frmNewPollutants"
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
' *  Subs:
' *     ExportLandCover(FileName as String) - creates text file of current classes
' *         Called By: mnuExpLCType
' *
' *  Misc:  Uses an invisible menu called mnuPopUp for right click events on the FlexGrid
' *************************************************************************************

Option Explicit

Private rsLCTypeCboClick As New ADODB.Recordset   'RS on cboPollName click event
Private rsCoeff As ADODB.Recordset
Private rsLCType As ADODB.Recordset
Private rsFullCoeff As ADODB.Recordset      'RS on cboCoeffSet click event
Private rsCoeffs As ADODB.Recordset         'RS that fills FlexGrid
Private rsWQStds As ADODB.Recordset         'Rs that fills FelexGrid on Standards Tab

Private boolLoaded As Boolean
Private boolChanged As Boolean              'Bool for enabling save button
Private boolDescChanged As Boolean          'Bool for seeing if Description Changed
Private boolSaved As Boolean                'Bool for whether or not things have saved

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

Private Sub cboLCType_Click()

    Dim strLCClasses As String
    Dim i As Integer
    
    strLCClasses = "SELECT LCTYPE.LCTYPEID, LCCLASS.VALUE, LCCLASS.NAME, LCCLASS.LCCLASSID FROM LCTYPE INNER JOIN LCCLASS ON " & _
    "LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'" & " ORDER BY LCCLASS.VALUE"
    
    rsLCTypeCboClick.Open strLCClasses, g_ADOConn, adOpenKeyset
    
    grdPollDef.Clear
    InitPollDefGrid grdPollDef
    grdPollDef.Rows = rsLCTypeCboClick.RecordCount + 1

    rsLCTypeCboClick.MoveFirst
    
    'Actually add the records to the new set
    For i = 1 To rsLCTypeCboClick.RecordCount

        With grdPollDef
            .TextMatrix(i, 1) = rsLCTypeCboClick!Value
            .TextMatrix(i, 2) = rsLCTypeCboClick!Name
            .TextMatrix(i, 3) = 0
            .TextMatrix(i, 4) = 0
            .TextMatrix(i, 5) = 0
            .TextMatrix(i, 6) = 0
            .TextMatrix(i, 7) = 0
            .TextMatrix(i, 8) = rsLCTypeCboClick!LCClassID
       
        rsLCTypeCboClick.MoveNext
        End With
        
    Next i
    
    rsLCTypeCboClick.Close
    Set rsLCTypeCboClick = Nothing
    
End Sub

Private Sub cmdSave_Click()

    If CheckForm Then
    
        If UpdateValues Then
            
            MsgBox txtPollutant.Text & " successfully added.  Please enter value for associated water quality standards.", vbInformation, "Pollutant Successfully Added"
            Unload Me
            frmPollutants.SSTab1.Tab = 1
        End If
    
    End If
    
End Sub

Private Sub Form_Load()
   
   Dim rsWQstd As New ADODB.Recordset
   Dim strSelectWQStd As String
   
   'Initialize the two flex Grids
   InitPollDefGrid grdPollDef
   InitPollWQStdGrid grdWQStd
   
   With frmNewPollutants
    .SSTab1.Tab = 0
    .SSTab1.TabEnabled(1) = False
    .SSTab1.TabVisible(1) = False
   End With
   
   modUtil.InitComboBox cboLCType, "LCType"

   boolLoaded = True
      
      
End Sub

Private Sub cmdQuit_Click()
   
   Unload Me

End Sub

Private Sub grdPollDef_Click()
    
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

                If IsNumeric(m_strUndoText) Then
                    txtActiveCell.Alignment = 1
                Else
                    txtActiveCell.Alignment = 0
                End If

             End If

        End With

    End If

    
End Sub

Private Sub grdPollDef_KeyDown(KeyCode As Integer, Shift As Integer)
    
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
    
    
End Sub

Private Sub grdWQStd_Click()
    
    Dim strUndoText As String
    
    txtActiveCellWQStd.Visible = False
    intRowWQ = grdWQStd.row
    intColWQ = grdWQStd.col

    'We want to limit editing to only the threshold column
    If intColWQ = 3 And intRowWQ >= 1 Then

        With txtActiveCellWQStd
            
            .Move SSTab1.Left + grdWQStd.Left + grdWQStd.CellLeft, _
            SSTab1.Top + grdWQStd.Top + grdWQStd.CellTop, _
            grdWQStd.CellWidth - 30, _
            grdWQStd.CellHeight - 150

            If intRowWQ <> 0 Then
                .Visible = True
                strUndoText = grdWQStd.TextMatrix(intRowWQ, intColWQ)
                .Text = grdWQStd.TextMatrix(intRowWQ, intColWQ)
                .SetFocus

                If IsNumeric(strUndoText) Then
                        txtActiveCellWQStd.Alignment = 1
                    Else
                        txtActiveCellWQStd.Alignment = 0
                End If

             End If

        End With

    End If
    
End Sub


Private Sub txtActiveCell_Change()
    
    grdPollDef.Text = txtActiveCell.Text
    boolChanged = True
    CmdSaveEnabled
    
End Sub

Private Sub txtActiveCell_GotFocus()

    txtActiveCell.SelStart = 0
    txtActiveCell.SelLength = Len(txtActiveCell.Text)

End Sub

Private Sub txtActiveCell_KeyDown(KeyCode As Integer, Shift As Integer)

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

End Sub

Private Sub KeyMoveUpdate()
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
                
                If IsNumeric(m_strUndoText) Then
                    txtActiveCell.Alignment = 1
                Else
                    txtActiveCell.Alignment = 0
                End If
                
            End If
        End With
        txtActiveCell.SetFocus
    
    ElseIf m_intPollCol = 0 Then
        
        txtActiveCell.Visible = False
    
    End If
    
End Sub

Private Sub txtActiveCell_Validate(Cancel As Boolean)

If Not boolChanged Then
    
    If txtActiveCell.Text <> m_strUndoText Then
        boolChanged = True
        CmdSaveEnabled
    End If

End If
End Sub

Private Sub txtActiveCellWQStd_Change()
    
    grdWQStd.Text = txtActiveCellWQStd.Text
    
End Sub

'Just making the entire cell text selected
Private Sub txtActiveCellWQStd_GotFocus()
    
    txtActiveCell.SelStart = 0
    txtActiveCellWQStd.SelLength = Len(txtActiveCellWQStd.Text)
    m_strUndoText = txtActiveCell.Text

End Sub

Private Sub mnuCoeffCopySet_Click()
    
    g_boolCopyCoeff = False
    Load frmCopyCoeffSet
    frmCopyCoeffSet.Show vbModal, Me

End Sub

Private Sub mnuCoeffNewSet_Click()
        
    g_boolAddCoeff = False
    Load frmAddCoeffSet
    frmAddCoeffSet.Show vbModal, Me
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Tab
        Case 0
            txtActiveCellWQStd.Visible = False
        Case 1
            txtActiveCell.Visible = False
    End Select

End Sub

'**********************************************************************************************************
'Subs/Functions
'**********************************************************************************************************
Public Sub CopyCoefficient(strNewCoeffName As String, strCoeffSet As String)

'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
'from the set being copied
    Dim strCopySet As String                'The Recordset of existing coefficients being copied
    Dim strLandClass As String              'Select for Landclass
    Dim i As Integer
    
    Dim rsCopySet As New ADODB.Recordset        'First RS
    Dim rsLandClass As New ADODB.Recordset      '1 record recordset to get a hold of LandClass
    
    strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & _
        "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
        
    rsCopySet.Open strCopySet, g_ADOConn, adOpenKeyset
    
    
    Debug.Print rsCopySet.RecordCount
    
    'Step 1: Enter name
    txtCoeffSet.Text = strNewCoeffName

    'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    grdPollDef.Clear
    grdPollDef.Rows = rsCopySet.RecordCount + 1

    rsCopySet.MoveFirst
    
    'Actually add the records to the new set
    For i = 1 To rsCopySet.RecordCount
        strLandClass = "SELECT * FROM LCCLASS WHERE LCCLASSID = " & rsCopySet!LCClassID
        'Let's try one more ADO method, why not, righ?
        rsLandClass.Open strLandClass, g_ADOConn, adOpenKeyset

        'Add the necessary components
        With grdPollDef
            .TextMatrix(i, 1) = rsLandClass!Value
            .TextMatrix(i, 2) = rsLandClass!Name
            .TextMatrix(i, 3) = rsCopySet!Coeff1
            .TextMatrix(i, 4) = rsCopySet!Coeff2
            .TextMatrix(i, 5) = rsCopySet!Coeff3
            .TextMatrix(i, 6) = rsCopySet!Coeff4
        
        End With
        rsLandClass.Close
        rsCopySet.MoveNext
    Next i
    
    
    'Set up everything to look good
    Unload frmCopyCoeffSet


    'Cleanup
    rsCopySet.Close
   
    Set rsCopySet = Nothing
    Set rsLandClass = Nothing
        
End Sub


Private Sub CmdSaveEnabled()
    If boolChanged Or boolDescChanged Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

'Grouped all of the checks here.
Private Function CheckForm() As Boolean
    
    If Trim(txtPollutant.Text) = "" Then
        MsgBox "Please enter a name for the new pollutant", vbCritical, "Enter Name"
        CheckForm = False
        txtPollutant.SetFocus
        txtPollutant.SelLength = Len(txtPollutant.Text)
        Exit Function
    ElseIf modUtil.UniqueName("Pollutant", txtPollutant) Then
        CheckForm = True
    Else
        MsgBox Err4, vbCritical, "Name In Use"
        CheckForm = False
        txtPollutant.SetFocus
        txtPollutant.SelLength = Len(txtPollutant.Text)
        Exit Function
    End If
    
    If Len(Trim(txtCoeffSet.Text)) = 0 Then
        MsgBox "Please enter a name for the new pollutant", vbCritical, "Enter Name"
        CheckForm = False
        txtPollutant.SetFocus
        txtPollutant.SelLength = Len(txtPollutant.Text)
        Exit Function
    ElseIf modUtil.UniqueName("Coefficientset", txtCoeffSet.Text) Then
        CheckForm = True
    Else
        MsgBox Err4, vbCritical, "Name In Use"
        CheckForm = False
        txtCoeffSet.SetFocus
        txtCoeffSet.SelLength = Len(txtPollutant.Text)
        Exit Function
    End If
    
    'Now if all is there and good, go on and check the grid values
    If ValidateGridValues Then
        CheckForm = True
    Else
        CheckForm = False
    End If
    
End Function


Private Function ValidateGridValues() As Boolean

'Need to validate each grid value before saving.  Essentially we take it a row at a time,
'then rifle through each column of each row.  Case Select tests each each x,y value depending
'on column... 3-6 must be 1-100 range

'Returns: True or False
    
    Dim varActive As Variant            'txtActiveCell value
    Dim i As Integer
    Dim j As Integer

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
        
    ValidateGridValues = True
        
End Function

Private Function UpdateValues() As Boolean

    Dim strInsertPollutant As String                'Insert String for new poll
    Dim strSelectPollutant As String                'Select string for new poll
    Dim strSelectLCType As String                   'Select string for LCType
    Dim strInsertCoeffSet As String                 'Insert string for new coeff set
    Dim strSelectCoeffSet As String                 'Select string for new coeff set
    Dim strInsertCoeffs As String                   'Insert string for the coefficients
    Dim strSelectWQCrit As String                   'Select string for Water Quality
    Dim strInsertWQCrit As String                   'Insert string for Water Quality
    Dim strNewColor As String                       'New color
    
    Dim rsNewPollutant As New ADODB.Recordset
    Dim rsNewCoeffSet As New ADODB.Recordset
    Dim rsLCType As New ADODB.Recordset
    Dim rsWQCrit As New ADODB.Recordset
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
On Error GoTo ErrHandler:
    'Step 1a: Get a new color for this pollutant
    strNewColor = modUtil.ReturnHSVColorString
        
    'Step 1: Insert the New Pollutant
    strInsertPollutant = "INSERT INTO POLLUTANT(NAME, POLLTYPE, COLOR) VALUES ('" & _
            Replace(Trim(txtPollutant.Text), "'", "''") & "', 0, " & "'" & _
            strNewColor & "'" & ")"
    
    g_ADOConn.Execute strInsertPollutant, adAffectCurrent
    
    'Step 2: Select the newly inserted pollutant info
    strSelectPollutant = "SELECT * FROM POLLUTANT WHERE NAME LIKE '" & Trim(txtPollutant.Text) & "'"
    rsNewPollutant.Open strSelectPollutant, g_ADOConn, adOpenDynamic
    
    'Step 2a: Select the WQ Standards
    strSelectWQCrit = "SELECT * FROM WQCriteria"
    rsWQCrit.Open strSelectWQCrit, g_ADOConn, adOpenKeyset
    
    rsWQCrit.MoveFirst
    
    For k = 1 To rsWQCrit.RecordCount
        
        strInsertWQCrit = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & _
                rsNewPollutant!POLLID & "," & _
                rsWQCrit!WQCRITID & "," & "0 )"
        g_ADOConn.Execute strInsertWQCrit, adAffectCurrent
            
        rsWQCrit.MoveNext
    Next k
    
    'Step 3: Get the LCtype information
    strSelectLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"
    rsLCType.Open strSelectLCType, g_ADOConn, adOpenDynamic
        
    'Step 4: Insert the New coefficient set
    strInsertCoeffSet = "INSERT INTO COEFFICIENTSET (NAME, DESCRIPTION, LCTYPEID, POLLID) VALUES (" & _
                        Replace(Trim(txtCoeffSet.Text), "'", "''") & "', '" & _
                        Replace(Trim(txtCoeffSetDesc.Text), "'", "''") & "'," & _
                        rsLCType!LCTypeID & "," & _
                        rsNewPollutant!POLLID & ")"
    g_ADOConn.Execute strInsertCoeffSet, adAffectCurrent
    
    'Step 5: Select the newly inserted coefficient set
    strSelectCoeffSet = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & txtCoeffSet.Text & "'"
    rsNewCoeffSet.Open strSelectCoeffSet, g_ADOConn, adOpenDynamic
    
    'Step 6: Insert the new coeffs for that set
    For i = 1 To grdPollDef.Rows - 1
        
        strInsertCoeffs = "INSERT INTO COEFFICIENT (COEFF1, COEFF2, COEFF3, COEFF4, COEFFSETID, LCCLASSID) VALUES (" & _
                grdPollDef.TextMatrix(i, 3) & ", " & _
                grdPollDef.TextMatrix(i, 4) & ", " & _
                grdPollDef.TextMatrix(i, 5) & ", " & _
                grdPollDef.TextMatrix(i, 6) & ", " & _
                rsNewCoeffSet!CoeffSetID & ", " & _
                grdPollDef.TextMatrix(i, 8) & ")"
        
        g_ADOConn.Execute strInsertCoeffs, adAffectCurrent
        
    Next i
    
    'Cleanup
    rsNewPollutant.Close
    rsLCType.Close
    rsNewCoeffSet.Close
    
    Set rsNewPollutant = Nothing
    Set rsLCType = Nothing
    Set rsNewCoeffSet = Nothing
    
    frmPollutants.cboPollName.Clear
    modUtil.InitComboBox frmPollutants.cboPollName, "Pollutant"
    frmPollutants.cboPollName.ListIndex = modUtil.GetCboIndex(txtPollutant.Text, frmPollutants.cboPollName)

    UpdateValues = True
    
Exit Function
ErrHandler:
    
    MsgBox "An error occurred while creating new pollutant." & vbNewLine & _
    Err.Number & ": " & Err.Description, vbCritical, "Error"
    
End Function

Public Sub AddCoefficient(strCoeffName As String, strLCType As String)

'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever
        
    Dim strNewLcType As String              'CmdString for inserting new coefficientset
    Dim strDefault As String                '
    Dim strNewCoeff As String
    Dim strNewCoeffID As String             'Holder for the CoefficientSetID
    Dim strInsertNewCoeff As String         'Putting the newly created coefficients in their table
    Dim intCoeffSetID As Integer
    Dim i As Integer
    
    Dim rsCopySet As ADODB.Recordset        'First RS
    Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    Dim rsNewCoeff As New ADODB.Recordset
    
    strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & _
            Replace(strCoeffName, "'", "''") & "'," & _
            m_intPollID & "," & _
            m_intLCTypeID & ")"
    
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
        "WHERE LCTYPE.Name Like " & "'" & strLCType & "'"
           
    Set rsCopySet = New ADODB.Recordset
    rsCopySet.Open strDefault, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    grdPollDef.Clear
    grdPollDef.Rows = rsCopySet.RecordCount + 1
    
    
    rsCopySet.MoveFirst
    modUtil.InitPollDefGrid grdPollDef  'Call this again to set it up as we cleared it
    
    Dim cmdInsert As New ADODB.Command
    
    'Now loopy loo to populate values.
    Dim strNewCoeff1 As String
    strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    
    rsNewCoeff.Open strNewCoeff1, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    For i = 1 To rsCopySet.RecordCount
                
        'Let's try one more ADO method, why not, righ?
        rsNewCoeff.AddNew
        
        'Add the necessary components
        rsNewCoeff!Coeff1 = 0
        rsNewCoeff!Coeff2 = 0
        rsNewCoeff!Coeff3 = 0
        rsNewCoeff!Coeff4 = 0
        rsNewCoeff!CoeffSetID = rsCoeffSetID!CoeffSetID
        rsNewCoeff!LCClassID = rsCopySet!LCClassID
        
       With grdPollDef
      
        .TextMatrix(i, 1) = rsCopySet!valValue
        .TextMatrix(i, 2) = rsCopySet!valName
        .TextMatrix(i, 3) = "0"
        .TextMatrix(i, 4) = "0"
        .TextMatrix(i, 5) = "0"
        .TextMatrix(i, 6) = "0"
        .TextMatrix(i, 7) = rsCoeffSetID!CoeffSetID
        .TextMatrix(i, 8) = rsNewCoeff!coeffID
        '.TextMatrix(i, 9) = rsNewCoeff!CoeffID
       
       End With
       rsCopySet.MoveNext
       
    Next i
    
    'Call the function to set everything to newly added Coefficient.
    'cboCoeffSet.ListIndex = GetCboIndex(strCoeffName, cboCoeffSet)
    
    'Cleanup
    rsCopySet.Close
    rsCoeffSetID.Close
   'rsNewCoeff.Close
    
    Set rsCopySet = Nothing
    Set rsCoeffSetID = Nothing
    Set rsNewCoeff = Nothing
    
    Unload Me
    
End Sub



