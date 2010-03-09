Imports System.Xml
Public Class clsXMLPollutantItems
    Implements System.Collections.IEnumerable
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLPollutantItems
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use a pollutants (the whole group) in the frmPrj Pollutants Tab
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


    Private Const NODE_NAME As String = "Pollutants"

    Private m_colItems As System.Collections.ArrayList

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



    Public ReadOnly Property Count() As Integer
        Get
            'Return the count of order items.

            Count = m_colItems.Count()

        End Get
    End Property


    Public Property Item(ByVal Index As Integer) As clsXMLPollutantItem
        Get
            'Get the order item at the given index.
            Item = m_colItems.Item(Index)
        End Get
        Set(ByVal Value As clsXMLPollutantItem)
            m_colItems.Add(Value)
        End Set
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = m_colItems.GetEnumerator
    End Function

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

        Dim clsMgmtScen As clsXMLPollutantItem

        For Each clsMgmtScen In m_colItems
            'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            node.appendChild(dom.CreateTextNode(vbNewLine & vbTab))
            clsMgmtScen.CreateNode(node)
        Next clsMgmtScen

        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.CreateTextNode(vbNewLine & vbTab))

        'UPGRADE_NOTE: Object clsMgmtScen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsMgmtScen = Nothing

        'Return the created node
        CreateNode = node

    End Function

    Public Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.

        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        Dim clsMgmtScen As clsXMLPollutantItem
        Dim nodes As XmlNodeList
        Dim MgmtNode As XmlNode

        clsMgmtScen = New clsXMLPollutantItem
        m_colItems = New Collections.ArrayList

        nodes = node.selectNodes(clsMgmtScen.NodeName)
        For Each MgmtNode In nodes
            clsMgmtScen = New clsXMLPollutantItem
            clsMgmtScen.LoadNode(MgmtNode)
            m_colItems.Add(clsMgmtScen)
        Next MgmtNode

        'UPGRADE_NOTE: Object clsMgmtScen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsMgmtScen = Nothing

    End Sub

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()

        m_colItems = New Collections.ArrayList

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()

        'UPGRADE_NOTE: Object m_colItems may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_colItems = Nothing

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

    'The NodeAppendAttribute subroutine creates an attribute having the given
    'name and value, and adds it to the given node's Attributes collection.
    Private Sub NodeAppendAttribute(ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal strName As String, ByVal varValue As Object)

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

        'Format
        'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        node.appendChild(dom.CreateTextNode(vbNewLine & vbTab))

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

    Public Sub Add(ByVal MgmtScen As clsXMLPollutantItem)
        'Add a mgmt scen item.

        m_colItems.Add(MgmtScen)

    End Sub

    Public Sub Remove(ByVal Index As Integer)
        'Remove an order item.

        m_colItems.Remove(Index)

    End Sub

    'End of code to support repeating <OrderItem> elements.
    '*************************************************************************
End Class