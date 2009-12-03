VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsHelp
' *************************************************************************************
' *  Description:  Class for Help menu option off NSPECT main menu
' *
' *
' *  Called By:  Help... menu
' *************************************************************************************
Implements ICommand
Private m_App As IApplication

'I threw the Ado.Close event here.  Keeps the Access dbase from locking.
Private Sub Class_Terminate()
    
On Error Resume Next:
    If Not (g_ADOConn Is Nothing) Then
        
        g_ADOConn.Close
    
    End If
    
    Set g_ADOConn = Nothing
    Set m_App = Nothing
    
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Help..."
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "N-SPECT"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Help"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "N-SPECTToolHelp"
End Property

Private Sub ICommand_OnClick()
On Error GoTo ErrHandler
    'Load the help file
    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "Introduction.htm"
Exit Sub

ErrHandler:
    MsgBox "Could not find NSPECT help.  Please check Help directory for NSPECT.chm.", vbCritical, "Help Not Found"
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set m_App = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Choose for help on N-SPECT"
End Property



