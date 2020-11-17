VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   6825
   ClientTop       =   6150
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtImpFile 
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2295
      Width           =   2760
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3195
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   1770
      TabIndex        =   0
      Top             =   1650
      Width           =   1230
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdBrowse_Click()
    
    'browse...get output filename
   dlgOpen.FileName = Empty
   With dlgOpen
     .Filter = MSG8
     .DialogTitle = MSG2
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   
   If Len(dlgOpen.FileName) > 0 Then
      txtImpFile.Text = Trim(dlgOpen.FileName)
      m_strFileName = txtImpFile.Text
      cmdOK.Enabled = True
   End If
   
   clsWrapperMain.XML = m_strFileName
   
   
   
   

End Sub





Private Sub Command1_Click()
     'browse...get output filename
   dlgOpen.FileName = Empty
   With dlgOpen
     .Filter = MSG8
     .DialogTitle = MSG2
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   
   If Len(dlgOpen.FileName) > 0 Then
      txtImpFile.Text = Trim(dlgOpen.FileName)
      m_strFileName = txtImpFile.Text
      'cmdOK.Enabled = True
   End If
   
   params.XML = m_strFileName
   
   
    
   
End Sub

Private Sub Form_Load()
    
    Set params = New clsWrapperMain
    
End Sub
