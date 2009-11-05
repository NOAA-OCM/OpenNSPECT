VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMainToolbar"
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

Option Explicit

Implements IToolBarDef
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "clsMainToolbar.cls"

Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long

Private Property Get IToolBarDef_Caption() As String
  On Error GoTo ErrorHandler
    
    IToolBarDef_Caption = "N-SPECT " & App.Major & "." & App.Minor & "." & App.Revision

  Exit Property
ErrorHandler:
  HandleError True, "IToolBarDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
  On Error GoTo ErrorHandler
    
    Select Case pos
        
        Case 0
            itemDef.id = "NSPECT.clsMnuProject"
            itemDef.Group = False
               
    End Select

  Exit Sub
ErrorHandler:
  HandleError True, "IToolBarDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get IToolBarDef_ItemCount() As Long
  On Error GoTo ErrorHandler

    IToolBarDef_ItemCount = 1

  Exit Property
ErrorHandler:
  HandleError True, "IToolBarDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get IToolBarDef_Name() As String
  On Error GoTo ErrorHandler

    IToolBarDef_Name = "N-SPECT"

  Exit Property
ErrorHandler:
  HandleError True, "IToolBarDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property


Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
  
    Dim fso As FileSystemObject
    Dim nspectPath As String
    Dim docPath As String
    
    Set fso = New FileSystemObject
    
    ' Detects and sets the path to N-SPECT's application folder (installation directory)
    nspectPath = App.path
    If Right(nspectPath, 1) = "\" Then
        nspectPath = Left(nspectPath, Len(nspectPath) - 1)
    End If
    If Right(nspectPath, 4) = "\bin" Then
        nspectPath = Left(nspectPath, Len(nspectPath) - 4)
    End If
    'MsgBox "N-SPECT path: " & nspectPath
    modUtil.g_nspectPath = nspectPath
    'SetEnvironmentVariable "NSPECTDAT", nspectPath
    'MsgBox "NSPECTDAT: " & Environ("NSPECTDAT")
    
    ' Detects and sets the path to N-SPECT's document folder (within user's "My Documents")
    'If Len(Environ("USERPROFILE")) = 0 Then
    '    Err.Raise vbObjectError + 100, "frmPrj.cmdOpenWS_Click", "Environment variable USERPROFILE is not available."
    'End If
    'docPath = Environ("USERPROFILE") & "\My Documents\" & App.CompanyName
    'If Not fso.FolderExists(docPath) Then
    '    MkDir docPath
    'End If
    'docPath = docPath & "\" & App.ProductName
    'If Not fso.FolderExists(docPath) Then
    '    MkDir docPath
    'End If
    'modUtil.g_nspectDocPath = docPath
    modUtil.g_nspectDocPath = nspectPath
        
    modUtil.ADOConnection
    
    Exit Sub
    
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
