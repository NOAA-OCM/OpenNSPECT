VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLenSlope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Define a class module for menu choice Precipitation...
Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Implements ICommand
Private m_App As IApplication



Private Sub Class_Terminate()
    Set m_App = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriCore.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Length/Slope GRID..."
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
    ICommand_Message = "Create Length/Slope GRID"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "N-SPECTToolLenSlope"
End Property

Private Sub ICommand_OnClick()
    ' initialize and show the form
    frmDialogHolder.init m_App
    frmDialogHolder.Show vbModal
  
'  With frmDialogHolder.dlgAMLOpen
'    .Filter = "Arc Macro Language File (*.aml)|*.aml"
'    .DialogTitle = "Choose the Length/Slope AML"
'    .DefaultExt = ".aml"
'    .InitDir = App.path
'    .ShowOpen
'  End With
'
'
'  If Len(frmDialogHolder.dlgAMLOpen.FileName) > 0 Then
'
'    Dim arcMod As ESRI.Arc
'    Dim arcResults As ESRIutil.Strings
'    Set arcResults = New ESRIutil.Strings
'
'    Set arcMod = New ESRI.Arc
'
'    Dim ArcStatus As Long
'    Dim cmdRunAML As String
'
'    cmdRunAML = "w " & modUtil.SplitWorkspaceName(frmDialogHolder.dlgAMLOpen.FileName)
'    ArcStatus = arcMod.Command(cmdRunAML, Nothing)
'    MsgBox ArcStatus
'
'    arcMod.PushString "test223"
'    arcMod.PushString "bubba"
'
'    cmdRunAML = "&run " & frmDialogHolder.dlgAMLOpen.FileName
'    ArcStatus = arcMod.Command(cmdRunAML, arcResults)
'
'    Dim varstring As String
'    varstring = arcMod.GetVariable("time")
'    MsgBox "The value of variable time: " & varstring
'
'  Else
'    Exit Sub
'  End If
   
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set m_App = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Create the Length/Slope GRID"
End Property

