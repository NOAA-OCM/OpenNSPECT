Imports System.Xml

Public Class clsXMLLUScenPollItem
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLLUScenPollItem
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use Pollutant Items with a management scenarios
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


    Private Const NODE_NAME As String = "LUScenPollutant"
    Private Const ATTRIBUTE_PollID As String = "ID"
    Private Const ELEMENT_Name As String = "PollName"
    Private Const ELEMENT_Type1 As String = "Type1"
    Private Const ELEMENT_Type2 As String = "Type2"
    Private Const ELEMENT_Type3 As String = "Type3"
    Private Const ELEMENT_Type4 As String = "Type4"

    Public intID As Short
    Public strPollName As String
    Public intType1 As Double
    Public intType2 As Double
    Public intType3 As Double
    Public intType4 As Double

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
            dom.LoadXml(Value)
            'UPGRADE_WARNING: Couldn't resolve default property of object dom.documentElement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LoadNode((dom.DocumentElement))

        End Set
    End Property

    Public Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
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

        NodeAppendAttribute(dom, node, ATTRIBUTE_PollID, intID)
        NodeAppendChildElement(dom, node, ELEMENT_Name, strPollName)
        NodeAppendChildElement(dom, node, ELEMENT_Type1, intType1)
        NodeAppendChildElement(dom, node, ELEMENT_Type2, intType2)
        NodeAppendChildElement(dom, node, ELEMENT_Type3, intType3)
        NodeAppendChildElement(dom, node, ELEMENT_Type4, intType4)

        'Return the created node
        CreateNode = node

    End Function

    Public Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.

        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        intID = CShort(GetNodeText(node, "@" & ATTRIBUTE_PollID))
        strPollName = GetNodeText(node, ELEMENT_Name)
        intType1 = CDbl(GetNodeText(node, ELEMENT_Type1))
        intType2 = CDbl(GetNodeText(node, ELEMENT_Type2))
        intType3 = CDbl(GetNodeText(node, ELEMENT_Type3))
        intType4 = CDbl(GetNodeText(node, ELEMENT_Type4))

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
        attr = dom.CreateAttribute(strName)
        'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        attr.Value = varValue

        node.Attributes.Append(attr)

    End Sub

    Private Sub NodeAppendChildElement(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal Name As String, ByVal Value As Object)
        'The NodeAppendChildElement subroutine creates an element having the given
        'name and value, and adds the element as a child of the given node.

        Dim element As XmlElement

        'Create a new child element and set its value.
        element = dom.CreateElement(Name)
        'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        element.InnerText = CStr(Value)

        'Append the new child element to the node.
        'UPGRADE_WARNING: Couldn't resolve default property of object element. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.AppendChild(element)

    End Sub

    Public Function GetNodeText(ByRef Parent As XmlNode, ByRef ChildName As String) As String
        'The GetNodeText function retrieves the value of the given child element
        'within the given parent element. If the requested child element is not
        'found, then an empty string is returned.

        Dim node As XmlNode

        On Error GoTo ErrorHandler

        GetNodeText = Parent.SelectSingleNode(ChildName).InnerText
        Exit Function

ErrorHandler:

        GetNodeText = ""

    End Function
End Class