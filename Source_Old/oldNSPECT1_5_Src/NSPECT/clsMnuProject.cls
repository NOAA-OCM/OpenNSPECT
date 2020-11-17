VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMnuProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsMnuProject
' *************************************************************************************
' *  Description: Class that defines the main menu in the N-Spect Toolbar
' *
' *
' *  Called By:
' *************************************************************************************
' *  Subs:
' *
' *
' *  Misc:
' *************************************************************************************

Option Explicit
' This class module implements the following interfaces
Implements esriSystemUI.IMenuDef         'Menu Definition
Implements esriFramework.IRootLevelMenu   'Root level so it can hang with the likes of File..., Edit...

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "clsMnuProject.cls"

Private Property Get IMenuDef_Caption() As String
  On Error GoTo ErrorHandler

    ' Menu caption
    IMenuDef_Caption = "N-SPECT"

  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

'The menu choices in the main menu of the Toolbar
Private Sub IMenuDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
  On Error GoTo ErrorHandler

On Err GoTo Errhand:
    ' Menu items
    Select Case pos
    Case 0
        itemDef.id = "NSPECT.clsNewAnalysis"
        itemDef.Group = False
    Case 1
        'Tool Setup -> (it's a shortcut menu)
        itemDef.id = "NSPECT.clsToolSetup"
        itemDef.Group = True
        'New Analysis...
    Case 2
        itemDef.id = "NSPECT.clsHelp"
        itemDef.Group = True
    End Select
Exit Sub
Errhand:
    MsgBox "Error with the IMenuDef in clsMnuProject", vbCritical, "error with MenuDef"

  Exit Sub
ErrorHandler:
  HandleError True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Property Get IMenuDef_ItemCount() As Long
  On Error GoTo ErrorHandler

    ' Menu count
    IMenuDef_ItemCount = 3

  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get IMenuDef_Name() As String
  On Error GoTo ErrorHandler

    'Menu name
    IMenuDef_Name = "clsMnuProject"

  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

