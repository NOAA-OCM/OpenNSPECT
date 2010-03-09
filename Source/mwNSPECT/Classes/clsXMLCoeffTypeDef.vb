Imports System.Xml

Public Class clsXMLCoeffTypeDef
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLCoeffTypeDef
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use Coefficient Type Definition
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
    Private Const NODE_NAME As String = "TypeDefFile"
    Private Const NODE_TDLyrName As String = "TDLyrName" 'Layer Name
    Private Const NODE_TDLyrFileName As String = "TDLyrFileName" 'Layer FileName
    Private Const NODE_TDAttribute As String = "TDAttribute"
    Private Const NODE_TDType As String = "TDType" 'Alpha/Numeric 0 = Alpha, 1 = Numeric
    Private Const NODE_TDDef1 As String = "TDDef1"
    Private Const NODE_TDDef2 As String = "TDDef2"
    Private Const NODE_TDDef3 As String = "TDDef3"
    Private Const NODE_TDDef4 As String = "TDDef4"

    'Variables holding value of nodes above
    Public strTDLyrName As String
    Public strTDLyrFileName As String
    Public strTDAttribute As String
    Public intTDType As Short
    Public strTDDef1 As String
    Public strTDDef2 As String
    Public strTDDef3 As String
    Public strTDDef4 As String

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
                dom.LoadXml(Value)
            End If

            node = dom.DocumentElement

            LoadNode(node)
        End Set
    End Property

    Public Sub SaveFile(ByRef strXML As String)

        Dim dom As New XmlDocument
        dom.LoadXml(strXML)

        'TODO
        'frmPrj.grdCoeffs.set_TextMatrix(g_intCoeffRow, 6, strXML)

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
            dom.LoadXml("<" & NODE_NAME & "/>")
            node = dom.DocumentElement
            'Otherwise use passed-in parent.
        Else
            dom = Parent.ownerDocument
            node = dom.CreateElement(NODE_NAME)
            Parent.appendChild(node)
        End If

        '*********************************************************************
        'TODO: Add code here to save attributes and child elements to the
        'node. Look to the commented-out code below for samples.
        NodeAppendChildElement(dom, node, NODE_TDLyrName, strTDLyrName)
        NodeAppendChildElement(dom, node, NODE_TDLyrFileName, strTDLyrFileName)
        NodeAppendChildElement(dom, node, NODE_TDAttribute, strTDAttribute)
        NodeAppendChildElement(dom, node, NODE_TDType, intTDType)
        NodeAppendChildElement(dom, node, NODE_TDDef1, strTDDef1)
        NodeAppendChildElement(dom, node, NODE_TDDef2, strTDDef2)
        NodeAppendChildElement(dom, node, NODE_TDDef3, strTDDef3)
        NodeAppendChildElement(dom, node, NODE_TDDef4, strTDDef4)

        'Return the created node

        CreateNode = node

    End Function


    Public Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.
        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        strTDLyrName = GetNodeText(node, NODE_TDLyrName)
        strTDLyrFileName = GetNodeText(node, NODE_TDLyrFileName)
        strTDAttribute = GetNodeText(node, NODE_TDAttribute)
        intTDType = CShort(GetNodeText(node, NODE_TDType))
        strTDDef1 = GetNodeText(node, NODE_TDDef1)
        strTDDef2 = GetNodeText(node, NODE_TDDef2)
        strTDDef3 = GetNodeText(node, NODE_TDDef3)
        strTDDef4 = GetNodeText(node, NODE_TDDef4)

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
        node.appendChild(element)

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