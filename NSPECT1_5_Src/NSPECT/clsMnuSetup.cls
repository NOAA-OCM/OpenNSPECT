VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMnuSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This class module implements the following interfaces
Implements IMenuDef
'Implements IRootLevelMenu

Private Property Get IMenuDef_Caption() As String
    ' Menu caption
    IMenuDef_Caption = "Tool Setup"
End Property

Private Sub IMenuDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriCore.IItemDef)
    ' Menu items
On Error GoTo Errhand:
    Select Case pos
    Case 0 'Land Cover Types
        itemDef.ID = "NSPECT.clsLCType"
        itemDef.Group = False
    Case 1 'Pollutants
        itemDef.ID = "NSPECT.clsPollutants"
        itemDef.Group = False
    Case 2 'Water Quality Standards
        itemDef.ID = "NSPECT.clsWQStd"
    Case 3 'Precipitation Scenarios
        itemDef.ID = "NSPECT.clsPrecip"
        itemDef.Group = False
    Case 4 'Delineate Watershed
        itemDef.ID = "NSPECT.clsWSDelin"
        itemDef.Group = False
    End Select
Exit Sub
Errhand:
    MsgBox "Error with IMenuDef in clsMnuSetup", vbCritical, "Error"
End Sub
Private Property Get IMenuDef_ItemCount() As Long
    ' Menu count
    IMenuDef_ItemCount = 5
End Property

Private Property Get IMenuDef_Name() As String
    'Menu name
    IMenuDef_Name = "clsMnuSetup"
End Property


