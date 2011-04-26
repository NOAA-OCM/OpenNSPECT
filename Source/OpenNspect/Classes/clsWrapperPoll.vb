'********************************************************************************************************
'File Name: clsWrapperPoll.vb
'Description: A class for handling the pollutant xml
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

Public Class clsXMLLUScenPollItem
    Inherits clsXMLBase
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

            NodeAppendAttribute(dom, node, ATTRIBUTE_PollID, intID)
            NodeAppendChildElement(dom, node, ELEMENT_Name, strPollName)
            NodeAppendChildElement(dom, node, ELEMENT_Type1, intType1)
            NodeAppendChildElement(dom, node, ELEMENT_Type2, intType2)
            NodeAppendChildElement(dom, node, ELEMENT_Type3, intType3)
            NodeAppendChildElement(dom, node, ELEMENT_Type4, intType4)

            'Return the created node
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
            If node Is Nothing Then Exit Sub

            intID = CShort(GetNodeText(node, "@" & ATTRIBUTE_PollID))
            strPollName = GetNodeText(node, ELEMENT_Name)
            intType1 = CDbl(GetNodeText(node, ELEMENT_Type1))
            intType2 = CDbl(GetNodeText(node, ELEMENT_Type2))
            intType3 = CDbl(GetNodeText(node, ELEMENT_Type3))
            intType4 = CDbl(GetNodeText(node, ELEMENT_Type4))

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class