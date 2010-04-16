Imports System.Xml

Public Class clsXMLLUScenPollItems
    Inherits clsXMLBase
    Implements System.Collections.IEnumerable
    '*************************************************************************
    'XML Wrapper Class
    '-------------------------------------------------------------------------
    'Use this template to create a class that is able to read and write XML
    'conforming to your specific schema. Refer to the MSDN Magazine article,
    '"Wrap Your XML Schema in Visual Basic Objects," for complete
    'documentation.
    '
    'Template Class Authors:
    '
    '    Dave Grundgeiger       daveg@tarasoftware.com
    '    Patrick Escarcega      patrick@vbguru.net
    '
    '*************************************************************************


    'The NODE_NAME constant contains the name of the XML element that
    'is being wrapped.
    Private Const NODE_NAME As String = "ManScenPollutants"

    Private m_colItems As Collections.ArrayList

    Public ReadOnly Property Item() As Collections.ArrayList
        Get
            Return m_colItems
        End Get
    End Property


    Public ReadOnly Property NodeName() As String
        Get
            'Retrieve the name of the element that this class wraps.

            NodeName = NODE_NAME

        End Get
    End Property


    Public ReadOnly Property Count() As Integer
        Get
            'Return the count of order items.
            Count = m_colItems.Count()

        End Get
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

        Dim clsPoll As clsXMLLUScenPollItem

        For Each clsPoll In m_colItems
            clsPoll.CreateNode(node)
        Next clsPoll

        'UPGRADE_NOTE: Object clsPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPoll = Nothing

        CreateNode = node

    End Function

    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.

        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub


        Dim clsPoll As clsXMLLUScenPollItem
        Dim nodes As XmlNodeList
        Dim PollNode As XmlNode

        clsPoll = New clsXMLLUScenPollItem
        m_colItems = New Collections.ArrayList

        nodes = node.SelectNodes(clsPoll.NodeName)
        For Each PollNode In nodes
            clsPoll = New clsXMLLUScenPollItem
            clsPoll.LoadNode(PollNode)
            m_colItems.Add(clsPoll)
        Next PollNode

        'UPGRADE_NOTE: Object clsPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsPoll = Nothing

    End Sub

    Public Sub New()
        m_colItems = New Collections.ArrayList
    End Sub



    Public Sub Add(ByVal Pollutant As clsXMLLUScenPollItem)
        'Add a pollutant item.
        m_colItems.Add(Pollutant)

    End Sub

    Public Sub Remove(ByVal Index As Integer)
        'Remove an order item.
        m_colItems.Remove(Index)

    End Sub

    'End of code to support repeating <OrderItem> elements.
    '*************************************************************************
End Class