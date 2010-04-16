Imports System.Xml

Public Class clsXMLBase
    Public Overridable Property XML() As String
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
            LoadNode((dom.DocumentElement))

        End Set
    End Property

    Public Overridable Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Return Nothing
    End Function

    Public Overridable Sub LoadNode(ByRef node As XmlNode)


    End Sub

    Public Function NodeGetText(ByVal node As XmlNode, ByVal strNodeName As String) As String
        'The NodeGetText function uses selectSingleNode on the passed-in node to
        'find the child node having the given name. When found, the text of that
        'node is returned. If the child node is not found in the node, then an
        'empty string is returned.

        NodeGetText = "" 'default return value
        On Error Resume Next
        NodeGetText = node.SelectSingleNode(strNodeName).Value

    End Function

    Public Sub NodeAppendAttribute(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal strName As String, ByVal varValue As Object)
        'The NodeAppendAttribute subroutine creates an attribute having the given
        'name and value, and adds it to the given node's Attributes collection.

        Dim attr As XmlAttribute

        'Create a new attribute and set its value.
        attr = dom.CreateAttribute(strName)
        attr.Value = varValue

        node.Attributes.Append(attr)

    End Sub

    public Sub NodeAppendChildElement(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal Name As String, ByVal Value As Object)
        'The NodeAppendChildElement subroutine creates an element having the given
        'name and value, and adds the element as a child of the given node.

        Dim element As XmlElement

        'Create a new child element and set its value.
        element = dom.CreateElement(Name)
        element.InnerText = CStr(Value)

        'Append the new child element to the node.
        node.AppendChild(element)
        'Format nicely
        node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

    End Sub

    Public Function GetNodeText(ByRef Parent As XmlNode, ByRef ChildName As String, Optional ByRef NodeTextType As String = "") As String
        'The GetNodeText function retrieves the value of the given child element
        'within the given parent element. If the requested child element is not
        'found, then an empty string is returned.

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
