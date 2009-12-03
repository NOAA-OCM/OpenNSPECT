VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportWQStd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Water Quality Standard"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4890
      Picture         =   "frmImportWQStd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3885
      TabIndex        =   4
      Top             =   1215
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2835
      TabIndex        =   3
      Top             =   1215
      Width           =   975
   End
   Begin VB.TextBox txtImpFile 
      Height          =   285
      Left            =   1845
      TabIndex        =   1
      Top             =   705
      Width           =   3000
   End
   Begin VB.TextBox txtStdName 
      Height          =   285
      Left            =   1845
      TabIndex        =   0
      Top             =   180
      Width           =   2430
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   330
      Top             =   1200
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Import File:"
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   6
      Top             =   750
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Standard Name:"
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "frmImportWQStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmImportWQStd
' *************************************************************************************
' *  Description:  Allows for the copying of the contents of an existing wat qual standard
' *
' *
' *  Called By:  frmWQStd
' *************************************************************************************

Option Explicit
Private m_strFileName As String
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmImportWQStd.frm"


Private Sub cmdBrowse_Click()
  On Error GoTo ErrorHandler

    
    'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = MSG1
     .DialogTitle = MSG2
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   If Len(dlgCMD1.FileName) > 0 Then
      txtImpFile.Text = Trim(dlgCMD1.FileName)
      m_strFileName = txtImpFile.Text
      cmdOK.Enabled = True
   End If

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

   Unload Me


  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler


    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
    Dim strLine As String
    Dim intLine As Integer
    Dim strName As String
    Dim strDescript As String
    Dim strPoll As String
    Dim strThresh As String
    Dim strCmd As String
    Dim rsNew As ADODB.Recordset
    
    Set fl = fso.OpenTextFile(m_strFileName, ForReading, True, TristateFalse)
    
    intLine = 0
    
    Do While Not fl.AtEndOfStream
        strLine = fl.ReadLine
        intLine = intLine + 1
        'MsgBox theLine
        
        If intLine = 1 Then
            
            strName = Trim(txtStdName.Text)
            strDescript = Split(strLine, ",")(1)
                             
            If strName = "" Then
            
                MsgBox "Name is blank.  Please enter a name.", vbCritical, "Empty Name Field"
                txtStdName.SetFocus
                Exit Sub
            
            Else
            
                strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & _
                        Replace(txtStdName.Text, "'", "''") & "', '" & _
                        Replace(strDescript, "'", "''") & "')"
                'Name Check
                    If modUtil.UniqueName("WQCRITERIA", txtStdName.Text) Then
                        g_ADOConn.Execute strCmd, adCmdText
                    Else
                        MsgBox "The name you have chosen is already in use.  Please select another.", vbCritical, "Select Unique Name"
                        Exit Sub
                    End If
                    
            End If
            
        Else
        
            strPoll = Split(strLine, ",")(0)
            strThresh = Split(strLine, ",")(1)
            'Insert the pollutant/threshold value into POLL_WQCRITERIA
            PollutantAdd strName, strPoll, strThresh
                
        End If
        
    Loop

    fl.Close
        
    'Cleanup
    frmWQStd.cboWQStdName.Clear
    modUtil.InitComboBox frmWQStd.cboWQStdName, "WQCRITERIA"
    frmWQStd.cboWQStdName.ListIndex = modUtil.GetCboIndex(txtStdName.Text, frmWQStd.cboWQStdName)
    
    Set fso = Nothing
    Set fl = Nothing
    Set rsNew = Nothing
    
    Unload Me
    
  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub PollutantAdd(strName As String, strPoll As String, strThresh As String)
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
               
    strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & _
            rsPollDetails!POLLID & "', '" & _
            rsPollAdd!WQCRITID & "'," & _
            strThresh & ")"
    
    g_ADOConn.Execute strCmdInsert, adCmdText
    
    'Cleanup
    rsPollAdd.Close
    rsPollDetails.Close
        
    Set rsPollAdd = Nothing
    Set rsPollDetails = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

    
