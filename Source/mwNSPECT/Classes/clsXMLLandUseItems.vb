'********************************************************************************************************
'File Name: clsXMLLandUseItems.vb
'Description: Class for handling land use items xml
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
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code

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
    Const c_sModuleFileName As String = "clsXMLLandUseItems.vb"


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
        Try
            GetEnumerator = m_colItems.GetEnumerator
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
            GetEnumerator = Nothing
        End Try
    End Function







    Public Overrides Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Try
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

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
            CreateNode = Nothing
        End Try
    End Function






    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        Try
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

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub





    Public Sub New()
        Try
            m_colItems = New Collections.ArrayList
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
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