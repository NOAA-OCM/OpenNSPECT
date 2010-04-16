Imports System.Xml

Public Class clsXMLLandUseItem
    Inherits clsXMLBase
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLLandUseItem
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use a single landuse scenario item
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


    Private Const NODE_NAME As String = "Landuse"
    Private Const ATTRIBUTE_ID As String = "ID"
    Private Const ELEMENT_Apply As String = "Apply"
    Private Const ELEMENT_LUScenName As String = "LUScenName"
    Private Const ELEMENT_LUScenXMLFile As String = "LUScenXMLFile"

    Public intID As Short
    Public intApply As Short
    Public strLUScenName As String
    Public strLUScenXMLFile As String

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

        node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
        NodeAppendAttribute(dom, node, ATTRIBUTE_ID, intID)
        NodeAppendChildElement(dom, node, ELEMENT_Apply, intApply)
        NodeAppendChildElement(dom, node, ELEMENT_LUScenName, strLUScenName)
        NodeAppendChildElement(dom, node, ELEMENT_LUScenXMLFile, strLUScenXMLFile)


        'Return the created node
        CreateNode = node

    End Function

    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.

        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        intID = CShort(GetNodeText(node, "@" & ATTRIBUTE_ID))
        intApply = CShort(GetNodeText(node, ELEMENT_Apply))
        strLUScenName = GetNodeText(node, ELEMENT_LUScenName)
        strLUScenXMLFile = GetNodeText(node, ELEMENT_LUScenXMLFile)

    End Sub

End Class