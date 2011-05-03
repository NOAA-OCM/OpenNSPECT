'********************************************************************************************************
'File Name: XmlLandUseItem.vb
'Description: Class for handling land use item xml
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

    Public Class LandUseItem
        Inherits Base
        ' *************************************************************************************
        ' *  Perot Systems Government Services
        ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
        ' *  XmlLandUseItem
        ' *************************************************************************************
        ' *  Description: Xml Wrapper for use a single landuse scenario item
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

        Private Const NODE_NAME As String = "Landuse"
        Private Const ATTRIBUTE_ID As String = "ID"
        Private Const ELEMENT_Apply As String = "Apply"
        Private Const ELEMENT_LUScenName As String = "LUScenName"
        Private Const ELEMENT_LUScenXmlFile As String = "LUScenXmlFile"

        Public intID As Short
        Public intApply As Short
        Public strLUScenName As String
        Public strLUScenXmlFile As String

        Public ReadOnly Property NodeName() As String
            Get
                'Retrieve the name of the element that this class wraps.

                NodeName = NODE_NAME

            End Get
        End Property

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

                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
                NodeAppendAttribute(dom, node, ATTRIBUTE_ID, intID)
                NodeAppendChildElement(dom, node, ELEMENT_Apply, intApply)
                NodeAppendChildElement(dom, node, ELEMENT_LUScenName, strLUScenName)
                NodeAppendChildElement(dom, node, ELEMENT_LUScenXmlFile, strLUScenXmlFile)

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
                If node Is Nothing Then Return

                intID = CShort(GetNodeText(node, "@" & ATTRIBUTE_ID))
                intApply = CShort(GetNodeText(node, ELEMENT_Apply))
                strLUScenName = GetNodeText(node, ELEMENT_LUScenName)
                strLUScenXmlFile = GetNodeText(node, ELEMENT_LUScenXmlFile)

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub
    End Class
End Namespace