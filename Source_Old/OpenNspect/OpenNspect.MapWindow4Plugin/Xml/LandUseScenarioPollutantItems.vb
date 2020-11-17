'********************************************************************************************************
'File Name: XmlPollItems
'Description: Class for handling pollutant items xml
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

Namespace Xml

    Public Class LandUseScenarioPollutantItems
        Inherits Base
        Implements IEnumerable
        '*************************************************************************
        'Xml Wrapper Class
        '-------------------------------------------------------------------------
        'Use this template to create a class that is able to read and write Xml
        'conforming to your specific schema. Refer to the MSDN Magazine article,
        '"Wrap Your Xml Schema in Visual Basic Objects," for complete
        'documentation. http://msdn.microsoft.com/en-us/magazine/cc302195.aspx
        '
        'Template Class Authors:
        '
        '    Dave Grundgeiger       daveg@tarasoftware.com
        '    Patrick Escarcega      patrick@vbguru.net
        '
        '*************************************************************************

        'The NODE_NAME constant contains the name of the Xml element that
        'is being wrapped.

        Private Const NODE_NAME As String = "ManScenPollutants"

        Private m_colItems As ArrayList

        Public ReadOnly Property Item() As ArrayList
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

        Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Try
                GetEnumerator = m_colItems.GetEnumerator
            Catch ex As Exception
                HandleError(ex)
                GetEnumerator = Nothing
            End Try
        End Function

        Public Overrides Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
            Try
                'Return an Xml DOM node that represents this class's properties. If a
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

                Dim Poll As LandUseScenarioPollutantItem

                For Each Poll In m_colItems
                    Poll.CreateNode(node)
                Next Poll

                CreateNode = node

            Catch ex As Exception
                HandleError(ex)
                CreateNode = Nothing
            End Try
        End Function

        Public Overrides Sub LoadNode(ByRef node As XmlNode)
            Try
                'Set this class's properties based on the data found in the
                'given node.

                'Ensure that a valid node was passed in.
                If node Is Nothing Then Return

                Dim Poll As LandUseScenarioPollutantItem
                Dim nodes As XmlNodeList
                Dim PollNode As XmlNode

                Poll = New LandUseScenarioPollutantItem
                m_colItems = New ArrayList

                nodes = node.SelectNodes(Poll.NodeName)
                For Each PollNode In nodes
                    Poll = New LandUseScenarioPollutantItem
                    Poll.LoadNode(PollNode)
                    m_colItems.Add(Poll)
                Next PollNode

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

        Public Sub New()
            Try
                m_colItems = New ArrayList
            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

        Public Sub Add(ByVal Pollutant As LandUseScenarioPollutantItem)
            Try
                'Add a pollutant item.
                m_colItems.Add(Pollutant)

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

        Public Sub Remove(ByVal Index As Integer)
            Try
                'Remove an order item.
                m_colItems.Remove(Index)

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

        'End of code to support repeating <OrderItem> elements.
        '*************************************************************************
    End Class
End Namespace