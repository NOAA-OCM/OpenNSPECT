VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWrapperMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private Const NODE_LCGridName = "LCGridName"
Private Const NODE_LCGridFileName = "LCGridFileName"
Private Const NODE_LCGridUnits = "LCGridUnits"
Private Const NODE_LCGridType = "LCGridType"
Private Const NODE_SoilsLayerName = "SoilsLayerName"
Private Const NODE_SoilsFileName = "SoilsFileName"
Private Const NODE_SoilsAttribute = "SoilsAttribute"
Private Const NODE_RainFallType = "RainFallType"
Private Const NODE_PrecipScenario = "PrecipScenario"
Private Const NODE_WaterShedDelin = "WaterShedDelineation"
Private Const NODE_WaterQuality = "WaterQualityStandard"
Private Const NODE_SelectedWaterSheds = "SelectedWaterSheds"
Private Const NODE_LocalEffects = "LocalEffects"
Private Const NODE_CalcErosion = "CalcErosion"
Private Const NODE_ErodAttribute = "ErodAttribute"
Private Const NODE_RainGridBool = "RainGridBool"
Private Const NODE_RainGridName = "RainGridName"
Private Const NODE_RainGridFileName = "RainGridFileName"
Private Const NODE_RainConstBool = "RainConstBool"
Private Const NODE_RainConstValue = "RainConstValue"


Private Const NODE_OutPutShapefile = "OutPutShapefile"
Private Const NODE_OutputLayerName = "OutputLayerName"

'Variables holding value of nodes above
Public strProjectName           As String
Public strLCGridName            As String
Public strLCGridFileName        As String
Public strLCGridUnits           As String
Public strLCGridType            As String
Public strSoilsLayerName        As String
Public strSoilsFileName         As String
Public strSoilsAttribute        As String
Public intRainFallType          As Integer
Public strPrecipScenario        As String
Public strWaterShedDelin        As String
Public strWaterQuality          As String
Public intSelectedWSheds        As Integer
Public intLocalEffects          As Integer

Public intCalcErosion           As Integer
Public strErodAttribute         As String
Public intRainGridBool          As Integer
Public strRainGridName          As String
Public strRainGridFileName      As String
Public intRainConstBool         As Integer
Public dblRainConstValue        As Double

Public strOutputShape           As String
Public strOutputLayer           As String



Public Sub LoadNode(node As MSXML2.IXMLDOMNode)

    Dim xmlMainNodeList As MSXML2.IXMLDOMNodeList
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlAttributeNode As MSXML2.IXMLDOMNode
    
    Set xmlMainNodeList = node.childNodes
        
    Dim intPollCount As Integer
    intPollCount = 0
    
    For Each xmlNode In xmlMainNodeList

        Select Case xmlNode.NodeName
    
        Case NODE_PRJNAME
            strProjectName = GetNodeText(node, NODE_PRJNAME)
        Case NODE_LCGridName
            strLCGridName = GetNodeText(node, NODE_LCGridName)
        Case NODE_LCGridFileName
            strLCGridFileName = GetNodeText(node, NODE_LCGridFileName) 'To Do, check
        Case NODE_LCGridUnits
            strLCGridUnits = GetNodeText(node, NODE_LCGridUnits)
        Case NODE_LCGridType
            strLCGridType = GetNodeText(node, NODE_LCGridType)
        Case NODE_SoilsLayerName
            strSoilsLayerName = GetNodeText(node, NODE_SoilsLayerName)
        Case NODE_SoilsFileName
            strSoilsFileName = GetNodeText(node, NODE_SoilsFileName)
        Case NODE_SoilsAttribute
            strSoilsAttribute = GetNodeText(node, NODE_SoilsAttribute)
        Case NODE_RainFallType
            intRainFallType = CInt(GetNodeText(node, NODE_RainFallType))
        Case NODE_PrecipScenario
            strPrecipScenario = GetNodeText(node, NODE_PrecipScenario)
        Case NODE_WaterShedDelin
            strWaterShedDelin = GetNodeText(node, NODE_WaterShedDelin)
        Case NODE_WaterQuality
            strWaterQuality = GetNodeText(node, NODE_WaterQuality)
        Case NODE_SelectedWaterSheds
            intSelectedWSheds = CInt(GetNodeText(node, NODE_SelectedWaterSheds))
        Case NODE_LocalEffects
            intLocalEffects = CInt(GetNodeText(node, NODE_LocalEffects))
'        Case NODE_Pollutant
'            strPollutant.Add GetNodeText(xmlNode, NODE_PollIgnore)
'            intPollIgnore.Add CInt(GetNodeText(xmlNode, NODE_PollIgnore)), CStr(intPollCount)
'            intPollCount = intPollCount + 1
        Case NODE_CalcErosion
            intCalcErosion = CInt(GetNodeText(node, NODE_CalcErosion))
        Case NODE_ErodAttribute
            strErodAttribute = GetNodeText(node, NODE_ErodAttribute)
        Case NODE_RainGridBool
            intRainGridBool = CInt(GetNodeText(node, NODE_RainGridBool))
        Case NODE_RainGridName
            strRainGridName = GetNodeText(node, NODE_RainGridName)
        Case NODE_RainGridFileName
            strRainGridFileName = GetNodeText(node, NODE_RainGridFileName)
        Case NODE_RainConstBool
            intRainConstBool = GetNodeText(node, NODE_RainConstBool)
        Case NODE_RainConstValue
            dblRainConstValue = CDbl(GetNodeText(node, NODE_RainConstValue))
        Case NODE_OutPutShapefile
            strOutputShape = GetNodeText(node, NODE_OutPutShapefile)
        Case NODE_OutputLayerName
            strOutputLayer = GetNodeText(node, NODE_OutputLayerName)
        End Select
        
    Next xmlNode
    
End Sub






'Retrieve the name of the element that this class wraps.
Public Property Get NodeName() As String
    NodeName = NODE_NAME
    
    If NodeName <> "ProjectFile" Then
        MsgBox "Error: The file you have choosen is not an N-SPECT project file.  Please select another", vbCritical, "Error Reading File"
    End If
    
End Property

'Retrieve the XML string that this class represents. The XML returned is
'built from the values of this class's properties.
Public Property Get XML() As String
    
    XML = Me.CreateNode().XML

End Property

'Loads the XML file from the CommonDialog selection
Public Property Let XML(ByVal strNewValue As String)
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load strNewValue

    dom.validateOnParse = True

    If dom.Validate.errorCode <> -1072897500 Then   'Error if associated DTD file is not present
        MsgBox dom.Validate.reason
    Else
    
    LoadNode dom.documentElement
    
    End If

End Property

Public Sub CreateNewXMLFile(strFileName As String)
    
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim xmlNode As MSXML2.IXMLDOMNode
    
    If xmlDoc.loadXML(g_strXMLTemplate) Then
        Set xmlNode = xmlDoc.documentElement
        NodeSetText xmlNode, NODE_PRJNAME, strProjectName
        NodeSetText xmlNode, NODE_LCGridName, strLCGridName
        NodeSetText xmlNode, NODE_LCGridFileName, strLCGridFileName
        NodeSetText xmlNode, NODE_LCGridUnits, strLCGridUnits
        NodeSetText xmlNode, NODE_LCGridType, strLCGridType
        NodeSetText xmlNode, NODE_SoilsLayerName, strSoilsLayerName
        NodeSetText xmlNode, NODE_SoilsFileName, strSoilsFileName
        NodeSetText xmlNode, NODE_SoilsAttribute, strSoilsAttribute
        NodeSetText xmlNode, NODE_RainFallType, CStr(intRainFallType)
        NodeSetText xmlNode, NODE_PrecipScenario, strPrecipScenario
        NodeSetText xmlNode, NODE_WaterShedDelin, strWaterShedDelin
        NodeSetText xmlNode, NODE_WaterQuality, strWaterQuality
        NodeSetText xmlNode, NODE_SelectedWaterSheds, CStr(intSelectedWSheds)
        NodeSetText xmlNode, NODE_LocalEffects, CStr(intLocalEffects)
        'TODO: Handle Pollutants, Scenarios
        
        
        
        NodeSetText xmlNode, NODE_CalcErosion, CStr(intCalcErosion)
        NodeSetText xmlNode, NODE_ErodAttribute, strErodAttribute
        NodeSetText xmlNode, NODE_RainGridBool, CStr(intRainGridBool)
        NodeSetText xmlNode, NODE_RainGridName, strRainGridName
        NodeSetText xmlNode, NODE_RainGridFileName, strRainGridFileName
        NodeSetText xmlNode, NODE_RainConstBool, CStr(intRainConstBool)
        NodeSetText xmlNode, NODE_RainConstValue, CStr(dblRainConstValue)
        NodeSetText xmlNode, NODE_OutPutShapefile, strOutputShape
        NodeSetText xmlNode, NODE_OutputLayerName, strOutputLayer
        
    End If
    
    xmlDoc.Save (strFileName)
        
End Sub


'Return an XML DOM node that represents this class's properties. If a
'parent DOM node is passed in, then the returned node is also added as a
'child node of the parent.
Public Function CreateNode(Optional Parent As MSXML2.IXMLDOMNode = Nothing) As MSXML2.IXMLDOMNode
    
    Dim node As MSXML2.IXMLDOMNode
    Dim dom As MSXML2.DOMDocument
    
    'If no parent was passed in, then create a DOM and document element.
    If Parent Is Nothing Then
        Set dom = New MSXML2.DOMDocument
        dom.loadXML g_strXMLTemplate
        Set node = dom.documentElement
    'Otherwise use passed-in parent.
    Else
        Set dom = Parent.ownerDocument
        Set node = dom.CreateElement(NODE_NAME)
        Parent.appendChild node
    End If
    
     NodeSetText node, NODE_PRJNAME, strProjectName
    

    Set CreateNode = node
    dom.Save ("C:\temp\xml.xml")
End Function

'The NodeGetText function uses selectSingleNode on the passed-in node to
'find the child node having the given name. When found, the text of that
'node is returned. If the child node is not found in the node, then an
'empty string is returned.
Private Function NodeGetText(ByVal node As MSXML2.IXMLDOMNode, ByVal strNodeName As String) As String
    NodeGetText = "" 'default return value
    On Error Resume Next
    NodeGetText = node.selectSingleNode(strNodeName).Text
End Function

'Sets value of single node
Private Function NodeSetText(ByVal node As MSXML2.IXMLDOMNode, ByVal strNodeName As String, strValue As String)
    
    node.selectSingleNode(strNodeName).Text = strValue
    
End Function

'The NodeAppendAttribute subroutine creates an attribute having the given
'name and value, and adds it to the given node's Attributes collection.
Private Sub NodeAppendAttribute( _
    ByVal dom As MSXML2.DOMDocument, _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal strName As String, _
    ByVal varValue As Variant _
    )
    
    Dim attr As MSXML2.IXMLDOMAttribute
    
    'Create a new attribute and set its value.
    Set attr = dom.createAttribute(strName)
    attr.Value = varValue
    
    'Add the new attribute to the node's Attributes collection.
    Set attr = node.Attributes.setNamedItem(attr)
    attr.nodeValue = varValue
    
End Sub

'The NodeAppendChildElement subroutine creates an element having the given
'name and value, and adds the element as a child of the given node.
Private Sub NodeAppendChildElement(ByVal dom As MSXML2.DOMDocument, ByVal node As MSXML2.IXMLDOMNode, _
    ByVal Name As String, ByVal Value As Variant)
    
    Dim element As MSXML2.IXMLDOMElement
    
    'Create a new child element and set its value.
    Set element = dom.getElementsByTagName(Name)
    element.Text = CStr(Value)

    'Append the new child element to the node.
    node.appendChild element
    
End Sub

'The GetNodeText function retrieves the value of the given child element
'within the given parent element. If the requested child element is not
'found, then an empty string is returned.
Public Function GetNodeText(Parent As MSXML2.IXMLDOMNode, ChildName As String) As String

    Dim node As MSXML2.IXMLDOMNode
    
    On Error GoTo ErrorHandler
    
        Set node = Parent.selectSingleNode(ChildName)
        GetNodeText = node.Text
        Debug.Print node.Text
    Exit Function
    
ErrorHandler:
    
    Debug.Print Err.Number & " "; Err.Description
    GetNodeText = ""
    
End Function



