VERSION 5.00
Begin VB.Form frmCopyWQStd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Water Quality Standard"
   ClientHeight    =   1725
   ClientLeft      =   7095
   ClientTop       =   6855
   ClientWidth     =   5505
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3945
      TabIndex        =   3
      Top             =   1215
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1215
      Width           =   975
   End
   Begin VB.TextBox txtStdName 
      Height          =   285
      Left            =   2625
      TabIndex        =   1
      Top             =   735
      Width           =   2370
   End
   Begin VB.ComboBox cboStdName 
      Height          =   315
      ItemData        =   "frmCopyWQStd.frx":0000
      Left            =   2625
      List            =   "frmCopyWQStd.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "New Standard Name:"
      Height          =   255
      Index           =   5
      Left            =   450
      TabIndex        =   5
      Top             =   765
      Width           =   1650
   End
   Begin VB.Label Label1 
      Caption         =   "Copy from Standard Name:"
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   4
      Top             =   255
      Width           =   2280
   End
End
Attribute VB_Name = "frmCopyWQStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmCopyWQStd
' *************************************************************************************
' *  Description:  Allows for the copying of the contents of an existing wat qual standard
' *
' *
' *  Called By:  frmWQStd
' *************************************************************************************

Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim strStandard As String
    Dim rsStandard As ADODB.Recordset
    
    Dim strPollStandard As String
    Dim rsPollStandard As ADODB.Recordset
    
    Dim rsNewStandard As ADODB.Recordset
    Dim rsNewPollCriteria As ADODB.Recordset
    
    Dim strCmd As String
    Dim strCmd2 As String
        
    'Get the WQ stand info
    strStandard = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboStdName.Text & "'"
   
    Set rsStandard = New ADODB.Recordset
    rsStandard.CursorLocation = adUseClient
    rsStandard.Open strStandard, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
      
    'Get the related pollutant/thresholds
    strPollStandard = "SELECT * FROM POLL_WQCRITERIA WHERE WQCRITID =" & rsStandard!WQCRITID
    
    Set rsPollStandard = New ADODB.Recordset
    rsPollStandard.CursorLocation = adUseClient
    rsPollStandard.Open strPollStandard, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic

    strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & _
            Replace(Trim(txtStdName.Text), "'", "''") & "', '" & _
            rsStandard!Description & "')"
      
    If modUtil.UniqueName("WQCRITERIA", Trim(txtStdName.Text)) Then
        Set rsNewStandard = modUtil.g_ADOConn.Execute(strCmd)
    Else
        MsgBox Err4, vbCritical, "Enter Unique Name"
        Exit Sub
    End If
    
    rsNewStandard.Open "Select * from WQCRITERIA WHERE NAME LIKE '" & Trim(txtStdName.Text) & "'", modUtil.g_ADOConn, adOpenDynamic
    
    Dim i As Integer
    i = 1
    
    rsPollStandard.MoveFirst
    
    For i = 1 To rsPollStandard.RecordCount
        strCmd2 = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & _
                rsPollStandard!POLLID & ", " & _
                rsNewStandard!WQCRITID & "," & _
                rsPollStandard!Threshold & ")"
         Set rsNewPollCriteria = modUtil.g_ADOConn.Execute(strCmd2)
         rsPollStandard.MoveNext
    Next i
      
    'Cleanup
    Set rsStandard = Nothing
    Set rsPollStandard = Nothing
    Set rsNewStandard = Nothing
    Set rsNewPollCriteria = Nothing
    
    frmWQStd.cboWQStdName.Clear
    modUtil.InitComboBox frmWQStd.cboWQStdName, "WQCRITERIA"
    
    Unload Me

End Sub

Private Sub PollutantAdd(strName As String, strPoll As String, strThresh As String)

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
               
    strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & _
            rsPollDetails!POLLID & "', '" & _
            rsPollAdd!WQCRITID & "'," & _
            strThresh & ")"
    
    g_ADOConn.Execute strCmdInsert, adCmdText
    
    Set rsPollAdd = Nothing
    Set rsPollDetails = Nothing
    
    
End Sub

Private Sub Form_Load()

    modUtil.InitComboBox cboStdName, "WQCRITERIA"
    
End Sub

