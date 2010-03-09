Imports System.Xml

Public Class clsXMLLUScen
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsWrapperMain
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use Management scenarios
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
    Private Const NODE_NAME As String = "LUScenFile"
    Private Const NODE_ManScenName As String = "LUScenName"
    Private Const NODE_ManScenLyrName As String = "LUScenLyrName"
    Private Const NODE_ManScenFileName As String = "LUScenFileName"
    Private Const NODE_ManScenSelectedPoly As String = "LUScenSelectedPoly"
    Private Const NODE_SCSCurveA As String = "SCSCurveA"
    Private Const NODE_SCSCurveB As String = "SCSCurveB"
    Private Const NODE_SCSCurveC As String = "SCSCurveC"
    Private Const NODE_SCSCurveD As String = "SCSCurveD"
    Private Const NODE_CoverFactor As String = "CoverFactor"
    Private Const NODE_WaterWetlands As String = "WaterWetlands"

    'Variables holding value of nodes above
    Public strLUScenName As String
    Public strLUScenLyrName As String
    Public strLUScenFileName As String
    Public intLUScenSelectedPoly As Short
    Public intSCSCurveA As Double
    Public intSCSCurveB As Double
    Public intSCSCurveC As Double
    Public intSCSCurveD As Double
    Public lngCoverFactor As Double
    Public intWaterWetlands As Short

    Public clsPollutant As clsXMLLUScenPollItem
    Public clsPollItems As clsXMLLUScenPollItems

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

            Dim dom As New xmlDocument
            Dim node As xmlNode

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

        Dim dom As New xmlDocument

        dom.loadXML(strXML)

        'TODO
        'frmPrj.grdLU.set_TextMatrix(CInt(g_intManScenRow), 3, strXML)

    End Sub

    Public Function CreateNode(Optional ByRef Parent As xmlNode = Nothing) As xmlNode
        'Return an XML DOM node that represents this class's properties. If a
        'parent DOM node is passed in, then the returned node is also added as a
        'child node of the parent.

        Dim node As xmlNode
        Dim dom As xmlDocument

        'If no parent was passed in, then create a DOM and document element.
        If Parent Is Nothing Then
            dom = New xmlDocument
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
        NodeAppendChildElement(dom, node, NODE_ManScenName, strLUScenName)
        NodeAppendChildElement(dom, node, NODE_ManScenLyrName, strLUScenLyrName)
        NodeAppendChildElement(dom, node, NODE_ManScenFileName, strLUScenFileName)
        NodeAppendChildElement(dom, node, NODE_ManScenSelectedPoly, intLUScenSelectedPoly)
        NodeAppendChildElement(dom, node, NODE_SCSCurveA, intSCSCurveA)
        NodeAppendChildElement(dom, node, NODE_SCSCurveB, intSCSCurveB)
        NodeAppendChildElement(dom, node, NODE_SCSCurveC, intSCSCurveC)
        NodeAppendChildElement(dom, node, NODE_SCSCurveD, intSCSCurveD)
        NodeAppendChildElement(dom, node, NODE_CoverFactor, lngCoverFactor)
        NodeAppendChildElement(dom, node, NODE_WaterWetlands, intWaterWetlands)

        clsPollItems.CreateNode(node)

        CreateNode = node

    End Function

    Public Sub LoadNode(ByRef node As xmlNode)
        'Set this class's properties based on the data found in the
        'given node.

        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        strLUScenName = GetNodeText(node, NODE_ManScenName)
        strLUScenLyrName = GetNodeText(node, NODE_ManScenLyrName)
        strLUScenFileName = GetNodeText(node, NODE_ManScenFileName)
        intLUScenSelectedPoly = CShort(GetNodeText(node, NODE_ManScenSelectedPoly))
        intSCSCurveA = CDbl(GetNodeText(node, NODE_SCSCurveA))
        intSCSCurveB = CDbl(GetNodeText(node, NODE_SCSCurveB))
        intSCSCurveC = CDbl(GetNodeText(node, NODE_SCSCurveC))
        intSCSCurveD = CDbl(GetNodeText(node, NODE_SCSCurveD))
        lngCoverFactor = CDbl(GetNodeText(node, NODE_CoverFactor))
        intWaterWetlands = CShort(GetNodeText(node, NODE_WaterWetlands))

        clsPollItems.LoadNode(node.selectSingleNode(clsPollItems.NodeName))

    End Sub

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()

        clsPollutant = New clsXMLLUScenPollItem
        clsPollItems = New clsXMLLUScenPollItems

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()

        'UPGRADE_NOTE: Object clsPollutant may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPollutant = Nothing
        'UPGRADE_NOTE: Object clsPollItems may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPollItems = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    Private Function NodeGetText(ByVal node As xmlNode, ByVal strNodeName As String) As String
        'The NodeGetText function uses selectSingleNode on the passed-in node to
        'find the child node having the given name. When found, the text of that
        'node is returned. If the child node is not found in the node, then an
        'empty string is returned.

        NodeGetText = "" 'default return value
        On Error Resume Next
        NodeGetText = node.SelectSingleNode(strNodeName).Value

    End Function

    Private Sub NodeAppendAttribute(ByVal dom As xmlDocument, ByVal node As xmlNode, ByVal strName As String, ByVal varValue As Object)
        'The NodeAppendAttribute subroutine creates an attribute having the given
        'name and value, and adds it to the given node's Attributes collection.

        Dim attr As XmlAttribute

        'Create a new attribute and set its value.
        attr = dom.createAttribute(strName)
        'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        attr.Value = varValue

        node.Attributes.Append(attr)

    End Sub

    Private Sub NodeAppendChildElement(ByVal dom As xmlDocument, ByVal node As xmlNode, ByVal Name As String, ByVal Value As Object)
        'The NodeAppendChildElement subroutine creates an element having the given
        'name and value, and adds the element as a child of the given node.

        Dim element As XmlElement

        'Create a new child element and set its value.
        element = dom.createElement(Name)
        'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        element.InnerText = CStr(Value)

        'Append the new child element to the node.
        'UPGRADE_WARNING: Couldn't resolve default property of object element. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(element)

    End Sub

    Public Function GetNodeText(ByRef Parent As xmlNode, ByRef ChildName As String) As String
        'The GetNodeText function retrieves the value of the given child element
        'within the given parent element. If the requested child element is not
        'found, then an empty string is returned.

        Dim node As xmlNode

        On Error GoTo ErrorHandler

        GetNodeText = Parent.SelectSingleNode(ChildName).InnerText
        Exit Function

ErrorHandler:

        GetNodeText = ""

    End Function
End Class