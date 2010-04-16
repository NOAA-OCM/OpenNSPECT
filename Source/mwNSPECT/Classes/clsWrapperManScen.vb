Imports System.Xml

Public Class clsXMLLUScen
    Inherits clsXMLBase
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


    Public Overrides Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        'Return an XML DOM node that represents this class's properties. If a
        'parent DOM node is passed in, then the returned node is also added as a
        'child node of the parent.

        Dim node As XmlNode
        Dim dom As XmlDocument

        'If no parent was passed in, then create a DOM and document element.
        If Parent Is Nothing Then
            dom = New XmlDocument
            dom.LoadXml("<" & NODE_NAME & "/>")
            node = dom.DocumentElement
            'Otherwise use passed-in parent.
        Else
            dom = Parent.OwnerDocument
            node = dom.CreateElement(NODE_NAME)
            Parent.AppendChild(node)
        End If

        '*********************************************************************
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

    Public Overrides Sub LoadNode(ByRef node As XmlNode)
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

        clsPollItems.LoadNode(node.SelectSingleNode(clsPollItems.NodeName))

    End Sub

    Public Sub New()
        clsPollutant = New clsXMLLUScenPollItem
        clsPollItems = New clsXMLLUScenPollItems

    End Sub

End Class