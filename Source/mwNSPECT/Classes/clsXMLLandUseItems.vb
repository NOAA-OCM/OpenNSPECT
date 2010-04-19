Imports System.Xml

Public Class clsXMLLandUseItems
    Inherits clsXMLBase
    Implements System.Collections.IEnumerable
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLLandUseItems
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use a landuse scenario items
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


    'The NODE_NAME constant contains the name of the XML element that
    'is being wrapped.
    Private Const NODE_NAME As String = "LandUses"

    Private m_colItems As Collections.ArrayList
    Public ReadOnly Property NodeName() As String
        Get
            'Retrieve the name of the element that this class wraps.

            NodeName = NODE_NAME

        End Get
    End Property



    'Return the count of order items.
    Public ReadOnly Property Count() As Integer
        Get
            Count = m_colItems.Count()
        End Get
    End Property

    Public Property Item(ByVal Index As Integer) As clsXMLLandUseItem
        Get
            'Get the order item at the given index.
            Item = m_colItems.Item(Index)
        End Get
        Set(ByVal Value As clsXMLLandUseItem)
            m_colItems.Add(Value)
        End Set
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = m_colItems.GetEnumerator
    End Function

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

        'Save the repeating <OrderItem> elements.

        Dim clsLandUse As clsXMLLandUseItem

        For Each clsLandUse In m_colItems
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
            clsLandUse.CreateNode(node)
        Next clsLandUse

        node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

        clsLandUse = Nothing

        CreateNode = node

    End Function

    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.


        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        Dim clsLandUse As clsXMLLandUseItem
        Dim nodes As XmlNodeList
        Dim LuNode As XmlNode

        clsLandUse = New clsXMLLandUseItem
        m_colItems = New Collections.ArrayList

        nodes = node.SelectNodes(clsLandUse.NodeName)
        For Each LuNode In nodes
            clsLandUse = New clsXMLLandUseItem
            clsLandUse.LoadNode(LuNode)
            m_colItems.Add(clsLandUse)
        Next LuNode

    End Sub

    Public Sub New()
        m_colItems = New Collections.ArrayList
    End Sub

    'Add an order item.
    Public Sub Add(ByVal LUItem As clsXMLLandUseItem)
        m_colItems.Add(LUItem)
    End Sub

    'Remove an order item.
    Public Sub Remove(ByVal Index As Integer)
        m_colItems.Remove(Index)
    End Sub
End Class