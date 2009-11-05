VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsToolSetup
' *************************************************************************************
' *  Description: Class that defines the tool setup sub-menu in the N-Spect menu
' *
' *
' *  Called By:
' *************************************************************************************
' *  Subs:
' *
' *  Misc:
' *************************************************************************************
Implements IMenuDef
Implements IShortcutMenu    'It's a shortcut menu

Private m_pApp As IApplication
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "clsToolSetup.cls"

Private Property Get IMenuDef_Caption() As String
  On Error GoTo ErrorHandler
    
    IMenuDef_Caption = "Advanced Settings"
    
  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub IMenuDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
  On Error GoTo ErrorHandler
     
   Select Case pos
    Case 0 'Land Cover Types
        itemDef.id = "NSPECT.clsLCType"
        itemDef.Group = False
    Case 1 'Pollutants
        itemDef.id = "NSPECT.clsPollutants"
        itemDef.Group = False
    Case 2 'Water Quality Standards
        itemDef.id = "NSPECT.clsWQStd"
    Case 3 'Precipitation Scenarios
        itemDef.id = "NSPECT.clsPrecip"
        itemDef.Group = False
    Case 4 'Delineate Watershed
        itemDef.id = "NSPECT.clsWSDelin"
        itemDef.Group = False
    Case 5 'Soils
        itemDef.id = "NSPECT.clsSoils"
        itemDef.Group = False
   End Select

  Exit Sub
ErrorHandler:
  HandleError True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get IMenuDef_ItemCount() As Long
  On Error GoTo ErrorHandler

    IMenuDef_ItemCount = 6

  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get IMenuDef_Name() As String
  On Error GoTo ErrorHandler
 
    IMenuDef_Name = "NSPECTMenu"
    
  Exit Property
ErrorHandler:
  HandleError True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
