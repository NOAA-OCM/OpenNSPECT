VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_papp As IApplication
Dim WithEvents p_mxdoc As MxDocument
Attribute p_mxdoc.VB_VarHelpID = -1
Implements IExtension
Implements IExtensionConfig   'Allows extension to appear in the Extensions dialog


Private m_extState As esriExtensionState

 
Private Property Get IExtension_Name() As String
  IExtension_Name = "NSPECT"
End Property
 
Private Sub IExtension_Startup(ByRef initializationData As Variant)

End Sub

 
Private Sub IExtension_Shutdown()
    
End Sub
 
Private Property Get IExtensionConfig_ProductName() As String
  IExtensionConfig_ProductName = "NSPECT"
End Property
 
Private Property Get IExtensionConfig_Description() As String
  IExtensionConfig_Description = "This enables the tools for this extension"
End Property
 
Private Property Get IExtensionConfig_State() As esriCore.esriExtensionState
   IExtensionConfig_State = m_extState
End Property
 
Private Property Let IExtensionConfig_State(ByVal ExtensionState As esriCore.esriExtensionState)
   m_extState = ExtensionState
End Property

