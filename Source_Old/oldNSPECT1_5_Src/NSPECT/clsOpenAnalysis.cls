VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpenAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Define a class module for menu choice Project...Open Analysis...
Implements ICommand
Private m_App As IApplication

Private Sub Class_Terminate()
    Set m_App = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriCore.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Open Analysis..."
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
    ICommand_Message = "Open Analysis"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "N-SPECTAnalysis"
End Property

Private Sub ICommand_OnClick()
    ' initialize and show the form
    frmPrj.init m_App
    frmPrj.Show vbModal
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set m_App = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Open Analysis"
End Property


