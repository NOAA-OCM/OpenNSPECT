VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLPrjFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsWrapperMain
' *************************************************************************************
' *  Description: XML Wrapper for use with main form's variables
' *
' *  Called By:
' *************************************************************************************
' *  Subs:
' *
' *
' *  Misc:
' *  Acknowledgements: Dave Grundgeiger       daveg@tarasoftware.com
' *                    Patrick Escarcega      patrick@vbguru.net
' *
' *************************************************************************************

Option Explicit

'Following are the names of the NODES
Private Const NODE_NAME = "NSPECTProjectFile"
Private Const NODE_PRJNAME = "PrjName"
Private Const NODE_PRJWORKSPACE = "PrjWorkspace"
Private Const NODE_LCGridName = "LCGridName"
Private Const NODE_LCGridFileName = "LCGridFileName"
Private Const NODE_LCGridUnits = "LCGridUnits"
Private Const NODE_LCGridType = "LCGridType"
Private Const NODE_SoilsDefName = "SoilsDefName"
Private Const NODE_SoilsHydFileName = "SoilsHydFileName"
Private Const NODE_SoilsKFileName = "SoilsKFileName"
Private Const NODE_PrecipScenario = "PrecipScenario"
Private Const NODE_WaterShedDelin = "WaterShedDelineation"
Private Const NODE_WaterQuality = "WaterQualityStandard"
Private Const NODE_SelectedPolys = "SelectedPolygons"
Private Const NODE_SelectedPolyLyrName = "SelectedPolyLyrName"
Private Const NODE_SelectedPolyFileName = "SelectedPolyFileName"
Private Const NODE_LocalEffects = "LocalEffects"
Private Const NODE_CalcErosion = "CalcErosion"
'****************** Added 12/03/07 **************************
Private Const NODE_UseOWNSDR = "SDRGrid"
Private Const NODE_SDRGridFileName = "SDRGridFileName"
'****************** end Add *********************************
Private Const NODE_RainGridBool = "RainGridBool"
Private Const NODE_RainGridName = "RainGridName"
Private Const NODE_RainGridFileName = "RainGridFileName"
Private Const NODE_RainConstBool = "RainConstBool"
Private Const NODE_RainConstValue = "RainConstValue"

'Variables holding value of nodes above
Public strProjectName           As String
Public strProjectWorkspace      As String
Public strLCGridName            As String
Public strLCGridFileName        As String
Public strLCGridUnits           As String
Public strLCGridType            As String
Public strSoilsDefName          As String
Public strSoilsHydFileName      As String
Public strSoilsKFileName        As String
Public strPrecipScenario        As String
Public strWaterShedDelin        As String
Public strWaterQuality          As String
Public intSelectedPolys         As Integer
Public strSelectedPolyFileName  As String
Public strSelectedPolyLyrName   As String
Public intLocalEffects          As Integer

'Class holders for DataGrid goodies
Public clsMgmtScenHolder As clsXMLMgmtScenItems     'A collection of management scenarios
Public clsMgmtScenItem As clsXMLMgmtScenItem        'A single Managment Scenario
Public clsPollItems As clsXMLPollutantItems         'A collection of pollutants from pollutants tab
Public clsPollItem As clsXMLPollutantItem           'A Single pollutant
Public clsLUItems As clsXMLLandUseItems             'A collection of land uses
Public clsLUItem As clsXMLLandUseItem               'A single land use

Public intCalcErosion           As Integer
Public intUseOwnSDR             As Integer
Public strSDRGridFileName       As String
Public intRainGridBool          As Integer
Public strRainGridName          As String
Public strRainGridFileName      As String
Public intRainConstBool         As Integer
Public dblRainConstValue        As Double

Public Property Get NodeName() As String
'Retrieve the name of the element that this class wraps.
    
    NodeName = NODE_NAME

End Property

Public Property Get XML() As String
'Retrieve the XML string that this class represents. The XML returned is
'built from the values of this class's properties.
    
    XML = Me.CreateNode().XML

End Property

Public Property Let XML(ByVal strNewValue As String)
'Assign a new XML string to this class. The newly assigned XML is parsed,
'and the class's properties are set accordingly.
    
    Dim dom As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    If InStr(strNewValue, ".xml") > 0 Then
        dom.Load strNewValue
    Else
        dom.loadXML strNewValue
    End If
   
    Set node = dom.documentElement
    
    LoadNode node

End Property

Public Sub SaveFile(strXML As String)
    
    Dim dom As New MSXML2.DOMDocument
    
    dom.loadXML Me.XML
    
    dom.Save strXML
    
End Sub

Public Function CreateNode(Optional Parent As MSXML2.IXMLDOMNode = Nothing) As MSXML2.IXMLDOMNode
'Return an XML DOM node that represents this class's properties. If a
'parent DOM node is passed in, then the returned node is also added as a
'child node of the parent.
    
    Dim node As MSXML2.IXMLDOMNode
    Dim dom As MSXML2.DOMDocument
    Dim comment As MSXML2.IXMLDOMComment
    
    'If no parent was passed in, then create a DOM and document element.
    If Parent Is Nothing Then
        Set dom = New MSXML2.DOMDocument
        dom.loadXML "<" & NODE_NAME & "/>"
        Set node = dom.documentElement
    'Otherwise use passed-in parent.
    Else
        Set dom = Parent.ownerDocument
        Set node = dom.createElement(NODE_NAME)
        Parent.appendChild node
    End If
    
    '*********************************************************************
    'TODO: Add code here to save attributes and child elements to the
    'node. Look to the commented-out code below for samples.
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    NodeAppendChildElement dom, node, NODE_PRJNAME, strProjectName
    NodeAppendChildElement dom, node, NODE_PRJWORKSPACE, strProjectWorkspace
    NodeAppendChildElement dom, node, NODE_LCGridName, strLCGridName
    NodeAppendChildElement dom, node, NODE_LCGridFileName, strLCGridFileName
    NodeAppendChildElement dom, node, NODE_LCGridUnits, strLCGridUnits
    NodeAppendChildElement dom, node, NODE_LCGridType, strLCGridType
    NodeAppendChildElement dom, node, NODE_SoilsDefName, strSoilsDefName
    NodeAppendChildElement dom, node, NODE_SoilsHydFileName, strSoilsHydFileName
    NodeAppendChildElement dom, node, NODE_SoilsKFileName, strSoilsKFileName
    'NodeAppendChildElement dom, node, NODE_RainFallType, intRainFallType
    NodeAppendChildElement dom, node, NODE_PrecipScenario, strPrecipScenario
    NodeAppendChildElement dom, node, NODE_WaterShedDelin, strWaterShedDelin
    NodeAppendChildElement dom, node, NODE_WaterQuality, strWaterQuality
    NodeAppendChildElement dom, node, NODE_SelectedPolys, intSelectedPolys
    NodeAppendChildElement dom, node, NODE_SelectedPolyFileName, strSelectedPolyFileName
    NodeAppendChildElement dom, node, NODE_SelectedPolyLyrName, strSelectedPolyLyrName
    NodeAppendChildElement dom, node, NODE_LocalEffects, intLocalEffects
    NodeAppendChildElement dom, node, NODE_CalcErosion, intCalcErosion
    NodeAppendChildElement dom, node, NODE_UseOWNSDR, intUseOwnSDR
    NodeAppendChildElement dom, node, NODE_SDRGridFileName, strSDRGridFileName
    NodeAppendChildElement dom, node, NODE_RainGridBool, intRainGridBool
    NodeAppendChildElement dom, node, NODE_RainGridName, strRainGridName
    NodeAppendChildElement dom, node, NODE_RainGridFileName, strRainGridFileName
    NodeAppendChildElement dom, node, NODE_RainConstBool, intRainConstBool
    NodeAppendChildElement dom, node, NODE_RainConstValue, dblRainConstValue
    
    'Format
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    
    'Pollutants
    clsPollItems.CreateNode node
    
    'Format
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    
    'Management Scenarios
    clsMgmtScenHolder.CreateNode node
    
    'Format
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    clsLUItems.CreateNode node
    
    'Format
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    
    Set CreateNode = node
    
End Function


Public Sub LoadNode(node As MSXML2.IXMLDOMNode)
'Set this class's properties based on the data found in the
'given node.

    'Ensure that a valid node was passed in.
    If node Is Nothing Then Exit Sub
    
    strProjectName = GetNodeText(node, NODE_PRJNAME)
    strProjectWorkspace = GetNodeText(node, NODE_PRJWORKSPACE)
    strLCGridName = GetNodeText(node, NODE_LCGridName)
    strLCGridFileName = GetNodeText(node, NODE_LCGridFileName)
    strLCGridUnits = GetNodeText(node, NODE_LCGridUnits)
    strLCGridType = GetNodeText(node, NODE_LCGridType)
    strSoilsDefName = GetNodeText(node, NODE_SoilsDefName)
    strSoilsHydFileName = GetNodeText(node, NODE_SoilsHydFileName)
    strSoilsKFileName = GetNodeText(node, NODE_SoilsKFileName)
    strPrecipScenario = GetNodeText(node, NODE_PrecipScenario)
    strWaterShedDelin = GetNodeText(node, NODE_WaterShedDelin)
    strWaterQuality = GetNodeText(node, NODE_WaterQuality)
    intSelectedPolys = CInt(GetNodeText(node, NODE_SelectedPolys))
    strSelectedPolyFileName = GetNodeText(node, NODE_SelectedPolyFileName)
    intLocalEffects = CInt(GetNodeText(node, NODE_LocalEffects))
    intCalcErosion = CInt(GetNodeText(node, NODE_CalcErosion))
    intUseOwnSDR = CInt(GetNodeText(node, NODE_UseOWNSDR, "integer"))
    strSDRGridFileName = GetNodeText(node, NODE_SDRGridFileName)
    intRainGridBool = CInt(GetNodeText(node, NODE_RainGridBool))
    strRainGridName = GetNodeText(node, NODE_RainGridName)
    strRainGridFileName = GetNodeText(node, NODE_RainGridFileName)
    intRainConstBool = GetNodeText(node, NODE_RainConstBool)
    dblRainConstValue = CDbl(GetNodeText(node, NODE_RainConstValue))
    
    clsMgmtScenHolder.LoadNode node.selectSingleNode(clsMgmtScenHolder.NodeName)
    clsPollItems.LoadNode node.selectSingleNode(clsPollItems.NodeName)
    clsLUItems.LoadNode node.selectSingleNode(clsLUItems.NodeName)
        
End Sub

Private Sub Class_Initialize()
    
    Set clsMgmtScenHolder = New clsXMLMgmtScenItems     'A collection of management scenarios
    Set clsMgmtScenItem = New clsXMLMgmtScenItem        'A single managment scenario
    Set clsPollItems = New clsXMLPollutantItems         'A collection of Pollutants
    Set clsPollItem = New clsXMLPollutantItem           'A single Pollutant
    Set clsLUItems = New clsXMLLandUseItems             'A collection of landuses
    Set clsLUItem = New clsXMLLandUseItem               'A Single LAnduse
    
End Sub

Private Sub Class_Terminate()

    Set clsMgmtScenHolder = Nothing
    Set clsMgmtScenItem = Nothing
    Set clsPollItems = Nothing
    Set clsPollItem = Nothing
    Set clsLUItems = Nothing
    Set clsLUItem = Nothing

End Sub

Private Function NodeGetText( _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal strNodeName As String) As String
'The NodeGetText function uses selectSingleNode on the passed-in node to
'find the child node having the given name. When found, the text of that
'node is returned. If the child node is not found in the node, then an
'empty string is returned.
        
    NodeGetText = "" 'default return value
    On Error Resume Next
    NodeGetText = node.selectSingleNode(strNodeName).Text

End Function

Private Sub NodeAppendAttribute( _
    ByVal dom As MSXML2.DOMDocument, _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal strName As String, _
    ByVal varValue As Variant)
'The NodeAppendAttribute subroutine creates an attribute having the given
'name and value, and adds it to the given node's Attributes collection.
    
    Dim attr As MSXML2.IXMLDOMAttribute
    
    'Create a new attribute and set its value.
    Set attr = dom.createAttribute(strName)
    attr.Value = varValue
    
    'Add the new attribute to the node's Attributes collection.
    Set attr = node.Attributes.setNamedItem(attr)
    attr.nodeValue = varValue
    
End Sub

Private Sub NodeAppendChildElement( _
    ByVal dom As MSXML2.DOMDocument, _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal Name As String, _
    ByVal Value As Variant)
'The NodeAppendChildElement subroutine creates an element having the given
'name and value, and adds the element as a child of the given node.
    
    Dim element As MSXML2.IXMLDOMElement
    
    'Create a new child element and set its value.
    Set element = dom.createElement(Name)
    element.Text = CStr(Value)

    'Append the new child element to the node.
    node.appendChild element
    'Format nicely
    node.appendChild dom.createTextNode(vbNewLine + vbTab)
    
End Sub

Public Function GetNodeText(Parent As MSXML2.IXMLDOMNode, ChildName As String, Optional NodeTextType As String) As String
'The GetNodeText function retrieves the value of the given child element
'within the given parent element. If the requested child element is not
'found, then an empty string is returned.

'Ed added optional parameter to return 0's for numerical values to avoid type mismatch for N-SPECT 1.5
'projects.

    Dim node As MSXML2.IXMLDOMNode
    
On Error GoTo ErrorHandler
    
    GetNodeText = Parent.selectSingleNode(ChildName).Text
    Exit Function
    
ErrorHandler:
    
    If NodeTextType = "integer" Then
        GetNodeText = "0"
    Else
        GetNodeText = ""
    End If
    
End Function
