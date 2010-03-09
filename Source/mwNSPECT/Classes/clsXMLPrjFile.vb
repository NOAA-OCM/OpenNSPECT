Imports System.Xml


Friend Class clsXMLPrjFile
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
	
	
	'Following are the names of the NODES
	Private Const NODE_NAME As String = "NSPECTProjectFile"
	Private Const NODE_PRJNAME As String = "PrjName"
	Private Const NODE_PRJWORKSPACE As String = "PrjWorkspace"
	Private Const NODE_LCGridName As String = "LCGridName"
	Private Const NODE_LCGridFileName As String = "LCGridFileName"
	Private Const NODE_LCGridUnits As String = "LCGridUnits"
	Private Const NODE_LCGridType As String = "LCGridType"
	Private Const NODE_SoilsDefName As String = "SoilsDefName"
	Private Const NODE_SoilsHydFileName As String = "SoilsHydFileName"
	Private Const NODE_SoilsKFileName As String = "SoilsKFileName"
	Private Const NODE_PrecipScenario As String = "PrecipScenario"
	Private Const NODE_WaterShedDelin As String = "WaterShedDelineation"
	Private Const NODE_WaterQuality As String = "WaterQualityStandard"
	Private Const NODE_SelectedPolys As String = "SelectedPolygons"
	Private Const NODE_SelectedPolyLyrName As String = "SelectedPolyLyrName"
	Private Const NODE_SelectedPolyFileName As String = "SelectedPolyFileName"
	Private Const NODE_LocalEffects As String = "LocalEffects"
	Private Const NODE_CalcErosion As String = "CalcErosion"
	'****************** Added 12/03/07 **************************
	Private Const NODE_UseOWNSDR As String = "SDRGrid"
	Private Const NODE_SDRGridFileName As String = "SDRGridFileName"
	'****************** end Add *********************************
	Private Const NODE_RainGridBool As String = "RainGridBool"
	Private Const NODE_RainGridName As String = "RainGridName"
	Private Const NODE_RainGridFileName As String = "RainGridFileName"
	Private Const NODE_RainConstBool As String = "RainConstBool"
	Private Const NODE_RainConstValue As String = "RainConstValue"
	
	'Variables holding value of nodes above
	Public strProjectName As String
	Public strProjectWorkspace As String
	Public strLCGridName As String
	Public strLCGridFileName As String
	Public strLCGridUnits As String
	Public strLCGridType As String
	Public strSoilsDefName As String
	Public strSoilsHydFileName As String
	Public strSoilsKFileName As String
	Public strPrecipScenario As String
	Public strWaterShedDelin As String
	Public strWaterQuality As String
	Public intSelectedPolys As Short
	Public strSelectedPolyFileName As String
	Public strSelectedPolyLyrName As String
	Public intLocalEffects As Short
	
	'Class holders for DataGrid goodies
	Public clsMgmtScenHolder As clsXMLMgmtScenItems 'A collection of management scenarios
	Public clsMgmtScenItem As clsXMLMgmtScenItem 'A single Managment Scenario
	Public clsPollItems As clsXMLPollutantItems 'A collection of pollutants from pollutants tab
	Public clsPollItem As clsXMLPollutantItem 'A Single pollutant
	Public clsLUItems As clsXMLLandUseItems 'A collection of land uses
	Public clsLUItem As clsXMLLandUseItem 'A single land use
	
	Public intCalcErosion As Short
	Public intUseOwnSDR As Short
	Public strSDRGridFileName As String
	Public intRainGridBool As Short
	Public strRainGridName As String
	Public strRainGridFileName As String
	Public intRainConstBool As Short
	Public dblRainConstValue As Double
	
	Public ReadOnly Property NodeName() As String
		Get
			'Retrieve the name of the element that this class wraps.
			
			NodeName = NODE_NAME
			
		End Get
	End Property
	
	
	Public Property XML() As String
		Get
			'Retrieve the XML string that this class represents. The XML returned is
			'built from the values of this class's properties.
			
            XML = Me.CreateNode().OuterXml
			
		End Get
		Set(ByVal Value As String)
			'Assign a new XML string to this class. The newly assigned XML is parsed,
			'and the class's properties are set accordingly.
			
            Dim dom As New XmlDocument
            Dim node As XmlNode

            If InStr(Value, ".xml") > 0 Then
                dom.Load(Value)
            Else
                dom.loadXML(Value)
            End If

            node = dom.documentElement

            LoadNode(node)

        End Set
    End Property

    Public Sub SaveFile(ByRef strXML As String)

        Dim dom As New XmlDocument
        dom.LoadXml(Me.XML)
        
        dom.Save(strXML)

    End Sub

    Public Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        'Return an XML DOM node that represents this class's properties. If a
        'parent DOM node is passed in, then the returned node is also added as a
        'child node of the parent.

        Dim node As XmlNode
        Dim dom As XmlDocument

        'If no parent was passed in, then create a DOM and document element.
        If Parent Is Nothing Then
            dom = New XmlDocument
            dom.loadXML("<" & NODE_NAME & "/>")
            node = dom.documentElement
            'Otherwise use passed-in parent.
        Else
            dom = Parent.ownerDocument
            node = dom.createElement(NODE_NAME)
            Parent.appendChild(node)
        End If

        '*********************************************************************
        'TODO: Add code here to save attributes and child elements to the
        'node. Look to the commented-out code below for samples.
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.createTextNode(vbNewLine & vbTab))
        NodeAppendChildElement(dom, node, NODE_PRJNAME, strProjectName)
        NodeAppendChildElement(dom, node, NODE_PRJWORKSPACE, strProjectWorkspace)
        NodeAppendChildElement(dom, node, NODE_LCGridName, strLCGridName)
        NodeAppendChildElement(dom, node, NODE_LCGridFileName, strLCGridFileName)
        NodeAppendChildElement(dom, node, NODE_LCGridUnits, strLCGridUnits)
        NodeAppendChildElement(dom, node, NODE_LCGridType, strLCGridType)
        NodeAppendChildElement(dom, node, NODE_SoilsDefName, strSoilsDefName)
        NodeAppendChildElement(dom, node, NODE_SoilsHydFileName, strSoilsHydFileName)
        NodeAppendChildElement(dom, node, NODE_SoilsKFileName, strSoilsKFileName)
        'NodeAppendChildElement dom, node, NODE_RainFallType, intRainFallType
        NodeAppendChildElement(dom, node, NODE_PrecipScenario, strPrecipScenario)
        NodeAppendChildElement(dom, node, NODE_WaterShedDelin, strWaterShedDelin)
        NodeAppendChildElement(dom, node, NODE_WaterQuality, strWaterQuality)
        NodeAppendChildElement(dom, node, NODE_SelectedPolys, intSelectedPolys)
        NodeAppendChildElement(dom, node, NODE_SelectedPolyFileName, strSelectedPolyFileName)
        NodeAppendChildElement(dom, node, NODE_SelectedPolyLyrName, strSelectedPolyLyrName)
        NodeAppendChildElement(dom, node, NODE_LocalEffects, intLocalEffects)
        NodeAppendChildElement(dom, node, NODE_CalcErosion, intCalcErosion)
        NodeAppendChildElement(dom, node, NODE_UseOWNSDR, intUseOwnSDR)
        NodeAppendChildElement(dom, node, NODE_SDRGridFileName, strSDRGridFileName)
        NodeAppendChildElement(dom, node, NODE_RainGridBool, intRainGridBool)
        NodeAppendChildElement(dom, node, NODE_RainGridName, strRainGridName)
        NodeAppendChildElement(dom, node, NODE_RainGridFileName, strRainGridFileName)
        NodeAppendChildElement(dom, node, NODE_RainConstBool, intRainConstBool)
        NodeAppendChildElement(dom, node, NODE_RainConstValue, dblRainConstValue)

        'Format
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.createTextNode(vbNewLine & vbTab))

        'Pollutants
        clsPollItems.CreateNode(node)

        'Format
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.createTextNode(vbNewLine & vbTab))

        'Management Scenarios
        clsMgmtScenHolder.CreateNode(node)

        'Format
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.createTextNode(vbNewLine & vbTab))
        clsLUItems.CreateNode(node)

        'Format
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.createTextNode(vbNewLine & vbTab))

        CreateNode = node

    End Function


    Public Sub LoadNode(ByRef node As XmlNode)
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
        intSelectedPolys = CShort(GetNodeText(node, NODE_SelectedPolys))
        strSelectedPolyFileName = GetNodeText(node, NODE_SelectedPolyFileName)
        intLocalEffects = CShort(GetNodeText(node, NODE_LocalEffects))
        intCalcErosion = CShort(GetNodeText(node, NODE_CalcErosion))
        intUseOwnSDR = CShort(GetNodeText(node, NODE_UseOWNSDR, "integer"))
        strSDRGridFileName = GetNodeText(node, NODE_SDRGridFileName)
        intRainGridBool = CShort(GetNodeText(node, NODE_RainGridBool))
        strRainGridName = GetNodeText(node, NODE_RainGridName)
        strRainGridFileName = GetNodeText(node, NODE_RainGridFileName)
        intRainConstBool = CShort(GetNodeText(node, NODE_RainConstBool))
        dblRainConstValue = CDbl(GetNodeText(node, NODE_RainConstValue))

        clsMgmtScenHolder.LoadNode(node.selectSingleNode(clsMgmtScenHolder.NodeName))
        clsPollItems.LoadNode(node.selectSingleNode(clsPollItems.NodeName))
        clsLUItems.LoadNode(node.selectSingleNode(clsLUItems.NodeName))

    End Sub

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()

        clsMgmtScenHolder = New clsXMLMgmtScenItems 'A collection of management scenarios
        clsMgmtScenItem = New clsXMLMgmtScenItem 'A single managment scenario
        clsPollItems = New clsXMLPollutantItems 'A collection of Pollutants
        clsPollItem = New clsXMLPollutantItem 'A single Pollutant
        clsLUItems = New clsXMLLandUseItems 'A collection of landuses
        clsLUItem = New clsXMLLandUseItem 'A Single LAnduse

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()

        'UPGRADE_NOTE: Object clsMgmtScenHolder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsMgmtScenHolder = Nothing
        'UPGRADE_NOTE: Object clsMgmtScenItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsMgmtScenItem = Nothing
        'UPGRADE_NOTE: Object clsPollItems may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPollItems = Nothing
        'UPGRADE_NOTE: Object clsPollItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPollItem = Nothing
        'UPGRADE_NOTE: Object clsLUItems may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsLUItems = Nothing
        'UPGRADE_NOTE: Object clsLUItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsLUItem = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    Private Function NodeGetText(ByVal node As XmlNode, ByVal strNodeName As String) As String
        'The NodeGetText function uses selectSingleNode on the passed-in node to
        'find the child node having the given name. When found, the text of that
        'node is returned. If the child node is not found in the node, then an
        'empty string is returned.

        NodeGetText = "" 'default return value
        On Error Resume Next
        NodeGetText = node.SelectSingleNode(strNodeName).Value

    End Function

    Private Sub NodeAppendAttribute(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal strName As String, ByVal varValue As Object)
        'The NodeAppendAttribute subroutine creates an attribute having the given
        'name and value, and adds it to the given node's Attributes collection.

        Dim attr As XmlAttribute

        'Create a new attribute and set its value.
        attr = dom.createAttribute(strName)
        'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        attr.Value = varValue

        node.Attributes.Append(attr)

    End Sub

    Private Sub NodeAppendChildElement(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal Name As String, ByVal Value As Object)
        'The NodeAppendChildElement subroutine creates an element having the given
        'name and value, and adds the element as a child of the given node.

        'Dim newnode As XmlNode = dom.CreateNode(XmlNodeType.)
        'newnode.Value = CStr(Value)
        'node.AppendChild(newnode)
        'node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

        Dim element As XmlElement

        'Create a new child element and set its value.
        element = dom.CreateElement(Name)
        'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        element.InnerText = CStr(Value)

        'Append the new child element to the node.
        'UPGRADE_WARNING: Couldn't resolve default property of object element. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.AppendChild(element)
        'Format nicely
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

    End Sub

    Public Function GetNodeText(ByRef Parent As XmlNode, ByRef ChildName As String, Optional ByRef NodeTextType As String = "") As String
        'The GetNodeText function retrieves the value of the given child element
        'within the given parent element. If the requested child element is not
        'found, then an empty string is returned.

        'Ed added optional parameter to return 0's for numerical values to avoid type mismatch for N-SPECT 1.5
        'projects.

        Dim node As XmlNode

        On Error GoTo ErrorHandler

        GetNodeText = Parent.SelectSingleNode(ChildName).InnerText
        Exit Function

ErrorHandler:

        If NodeTextType = "integer" Then
            GetNodeText = "0"
        Else
            GetNodeText = ""
        End If

    End Function
End Class