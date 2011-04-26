'********************************************************************************************************
'File Name: clsXMLOutputItems.vb
'Description: Class for handling Output items xml
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Dec 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.Xml

Public Class clsXMLOutputItems
    Inherits clsXMLBase
    Implements IEnumerable

    'The NODE_NAME constant contains the name of the XML element that
    'is being wrapped.
    Private Const NODE_NAME As String = "OutputFiles"

    Private m_colItems As ArrayList

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

    Public Property Item (ByVal Index As Integer) As clsXMLOutputItem
        Get
            'Get the order item at the given index.
            Item = m_colItems.Item (Index)
        End Get
        Set (ByVal Value As clsXMLOutputItem)
            m_colItems.Add (Value)
        End Set
    End Property

    Public Function GetEnumerator() As IEnumerator _
        Implements IEnumerable.GetEnumerator
        Try
            GetEnumerator = m_colItems.GetEnumerator
        Catch ex As Exception
            HandleError (ex)
            GetEnumerator = Nothing
        End Try
    End Function

    Public Overrides Function CreateNode (Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Try
            'Return an XML DOM node that represents this class's properties. If a
            'parent DOM node is passed in, then the returned node is also added as a
            'child node of the parent.

            Dim node As XmlNode
            Dim dom As XmlDocument

            'If no parent was passed in, then create a DOM and document element.
            If Parent Is Nothing Then
                dom = New XmlDocument
                dom.LoadXml ("<" & NODE_NAME & "/>")
                node = dom.DocumentElement
                'Otherwise use passed-in parent.
            Else
                dom = Parent.OwnerDocument
                node = dom.CreateElement (NODE_NAME)
                Parent.AppendChild (node)
            End If

            'Save the repeating <OrderItem> elements.

            Dim clsOutput As clsXMLOutputItem

            For Each clsOutput In m_colItems
                node.AppendChild (dom.CreateTextNode (vbNewLine & vbTab))
                clsOutput.CreateNode (node)
            Next clsOutput

            node.AppendChild (dom.CreateTextNode (vbNewLine & vbTab))

            clsOutput = Nothing

            CreateNode = node

        Catch ex As Exception
            HandleError (ex)
            CreateNode = Nothing
        End Try
    End Function

    Public Overrides Sub LoadNode (ByRef node As XmlNode)
        Try
            'Set this class's properties based on the data found in the
            'given node.

            'Ensure that a valid node was passed in.
            If node Is Nothing Then Exit Sub

            Dim clsOutput As clsXMLOutputItem
            Dim nodes As XmlNodeList
            Dim outNode As XmlNode

            clsOutput = New clsXMLOutputItem
            m_colItems = New ArrayList

            nodes = node.SelectNodes (clsOutput.NodeName)
            For Each outNode In nodes
                clsOutput = New clsXMLOutputItem
                clsOutput.LoadNode (outNode)
                m_colItems.Add (clsOutput)
            Next outNode

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Public Sub New()
        Try
            m_colItems = New ArrayList
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    'Add an order item.
    Public Sub Add (ByVal OutItem As clsXMLOutputItem)
        m_colItems.Add (OutItem)
    End Sub

    'Remove an order item.
    Public Sub Remove (ByVal Index As Integer)
        m_colItems.Remove (Index)
    End Sub
End Class