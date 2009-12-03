VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportLCType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Land Cover Type"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLCType 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   255
      Width           =   2370
   End
   Begin VB.TextBox txtImpFile 
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   735
      Width           =   3540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2850
      TabIndex        =   3
      Top             =   1215
      Width           =   976
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3885
      TabIndex        =   4
      Top             =   1215
      Width           =   976
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4965
      Picture         =   "frmImportLCType.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   735
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   255
      Top             =   1260
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "New Land Cover Type Name:"
      Height          =   255
      Index           =   5
      Left            =   195
      TabIndex        =   6
      Top             =   300
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Import File:"
      Height          =   255
      Index           =   0
      Left            =   225
      TabIndex        =   5
      Top             =   750
      Width           =   1335
   End
End
Attribute VB_Name = "frmImportLCType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmNewPrecip
' *************************************************************************************
' *  Description:  Allows for the importing of an existing land cover type classification
' *
' *
' *  Called By:  frmImportLCType
' *************************************************************************************

Option Explicit

Private m_strFileName As String
Private booName As Boolean  'Check if user put a name in
Private booFile As Boolean  'Check if FileName is correct


Private Sub cmdBrowse_Click()
   'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = Replace(MSG1, "<name>", "Land Cover Type")
     .DialogTitle = Replace(MSG2, "<name>", "Land Cover Type")
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   
   If Len(dlgCMD1.FileName) > 0 Then
      txtImpFile.Text = Trim(dlgCMD1.FileName)
      m_strFileName = Trim(dlgCMD1.FileName)
      booFile = True
   End If
   
   If booFile And booName Then
    cmdOK.Enabled = True
   End If

End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
    Dim strLine As String
    Dim intLine As Integer
    Dim strName As String
    Dim strDescript As String
    Dim strParams(7) As Variant
    Dim strCmd As String
    Dim rsNew As ADODB.Recordset
    Dim i As Integer
    
    If fso.FileExists(m_strFileName) Then
    
        Set fl = fso.OpenTextFile(m_strFileName, ForReading, True, TristateFalse)
        
        intLine = 0
        
        Do While Not fl.AtEndOfStream
            strLine = fl.ReadLine
            intLine = intLine + 1
            'MsgBox theLine
            
            'Get the first line, supposed to contain Name, Description
            If intLine = 1 Then
                
                strName = Trim(txtLCType.Text)
                strDescript = Split(strLine, ",")(1)
                                 
                'Check if name is present, if not bark
                If strName = "" Then
                
                    MsgBox "Name is blank.  Please enter a name.", vbCritical, "Empty Name Field"
                    txtLCType.SetFocus
                    Exit Sub
                
                Else
                    
                    'Name Check, if cool perform
                    If modUtil.UniqueName("LCTYPE", txtLCType.Text) Then
                        strCmd = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & _
                                Replace(txtLCType.Text, "'", "''") & "', '" & _
                                Replace(strDescript, "'", "''") & "')"
                        g_ADOConn.Execute strCmd, adCmdText
                    Else
                        MsgBox "The name you have chosen is already in use.  Please select another.", vbCritical, "Select Unique Name"
                        Exit Sub
                    End If 'End unique name check
                        
                End If 'end empty name check
                
            Else  ' > line 1
                
                If Len(Trim(strLine)) > 0 Then
                
                    i = 0
                    
                    'Create an array of lines ie Value,Descript,1,2,3,4,CoverFactor,W/WL
                    For i = 0 To UBound(strParams)
                        strParams(i) = Split(strLine, ",")(i)
                    Next i
                    
                    'Check the values, if ok add them, if not rollback
                    If modLCTypeValues.CheckGridValuesLCType(strParams) Then
                        AddLCClass strName, strParams
                    Else
                        modLCTypeValues.RollBack (strName)
                        GoTo Cleanup
                    End If 'End check
                Else
                    GoTo Cleanup
                End If
                    
            End If
            
        Loop
    
        fl.Close
        'Redo the form and make newly added lctype first
        frmLCTypes.cboLCType.Clear
        modUtil.InitComboBox frmLCTypes.cboLCType, "LCType"
        frmLCTypes.cboLCType.ListIndex = modUtil.GetCboIndex(strName, frmLCTypes.cboLCType)
        Unload Me
    Else
        MsgBox "The file you are pointing to does not exist. Please select another.", vbCritical, "File Not Found"
    'Cleanup
    End If
    
Exit Sub
    
Cleanup:
    frmLCTypes.cboLCType.Clear
    modUtil.InitComboBox frmLCTypes.cboLCType, "LCTYPE"
            
    Set fso = Nothing
    Set fl = Nothing
    Set rsNew = Nothing
    
    Unload Me
    
    
End Sub

Private Sub AddLCClass(strName As String, strParams())

On Error GoTo ErrHandler

    Dim strLCTypeAdd As String
    Dim strCmdInsert As String
    
    Dim rsLCType As ADODB.Recordset
    Set rsLCType = New ADODB.Recordset
    
       
    'Get the WQCriteria values using the name
    strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
    rsLCType.CursorLocation = adUseClient
    rsLCType.Open strLCTypeAdd, g_ADOConn, adOpenDynamic, adLockOptimistic
               
    'On Error GoTo dbaseerr
    
    strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & _
            Replace(CStr(strParams(0)), "'", "''") & ",'" & _
            Replace(CStr(strParams(1)), "'", "''") & "'," & _
            Replace(CStr(rsLCType!LCTypeID), "'", "''") & "," & _
            Replace(CStr(strParams(2)), "'", "''") & "," & _
            Replace(CStr(strParams(3)), "'", "''") & "," & _
            Replace(CStr(strParams(4)), "'", "''") & "," & _
            Replace(CStr(strParams(5)), "'", "''") & "," & _
            Replace(CStr(strParams(6)), "'", "''") & "," & _
            Replace(CStr(strParams(7)), "'", "''") & ")"
  
    g_ADOConn.Execute strCmdInsert, adCmdText
    
    Set rsLCType = Nothing

    Exit Sub
    
ErrHandler:
    
    MsgBox "There was a problem updating the database.  Insure that your values meet the correct " & _
    "value ranges for each field.", vbCritical, "Invalid Values Found"
    

    
End Sub




Private Sub txtLCType_Change()

    If Trim(Len(txtLCType.Text)) Then
        booName = True
    End If
    
    If booFile And booName Then
        cmdOK.Enabled = True
    End If
    
End Sub
