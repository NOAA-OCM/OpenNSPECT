Imports System.Xml

Public Class clsXMLPollutantItem
    Inherits clsXMLBase
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLPollutantItem
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use a single pollutant in the frmPrj Pollutants Tab
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


    Private Const NODE_NAME As String = "Pollutant"
    Private Const ATTRIBUTE_ID As String = "ID"
    Private Const ELEMENT_Apply As String = "Apply"
    Private Const ELEMENT_PollName As String = "PollName"
    Private Const ELEMENT_CoeffSet As String = "CoeffSet"
    Private Const ELEMENT_Coeff As String = "Coeff"
    Private Const ELEMENT_Threshold As String = "Threshold"
    Private Const ELEMENT_TypeDefXMLFile As String = "TypeDefXMLFile"

    Public intID As Short
    Public intApply As Short
    Public strPollName As String
    Public strCoeffSet As String
    Public strCoeff As String
    Public intThreshold As Short
    Public strTypeDefXMLFile As String

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
        NodeAppendChildElement(dom, node, ELEMENT_PollName, strPollName)
        NodeAppendChildElement(dom, node, ELEMENT_CoeffSet, strCoeffSet)
        NodeAppendChildElement(dom, node, ELEMENT_Coeff, strCoeff)
        NodeAppendChildElement(dom, node, ELEMENT_Threshold, intThreshold)
        NodeAppendChildElement(dom, node, ELEMENT_TypeDefXMLFile, strTypeDefXMLFile)
        node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

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
        strPollName = GetNodeText(node, ELEMENT_PollName)
        strCoeffSet = GetNodeText(node, ELEMENT_CoeffSet)
        strCoeff = GetNodeText(node, ELEMENT_Coeff)
        intThreshold = CShort(GetNodeText(node, ELEMENT_Threshold))
        strTypeDefXMLFile = GetNodeText(node, ELEMENT_TypeDefXMLFile)

    End Sub
End Class