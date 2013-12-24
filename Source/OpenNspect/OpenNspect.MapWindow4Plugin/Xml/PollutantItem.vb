'********************************************************************************************************
'File Name: XmlPollutantItem.vb
'Description: Class for handling a pollutant item xml
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

    Public Class PollutantItem
        Inherits Base
        ' *************************************************************************************
        ' *  Perot Systems Government Services
        ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
        ' *  XmlPollutantItem
        ' *************************************************************************************
        ' *  Description: Xml Wrapper for use a single pollutant in the frmPrj Pollutants Tab
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

        Private Const NODE_NAME As String = "Pollutant"
        Private Const ATTRIBUTE_ID As String = "ID"
        Private Const ELEMENT_Apply As String = "Apply"
        Private Const ELEMENT_PollName As String = "PollName"
        Private Const ELEMENT_CoeffSet As String = "CoeffSet"
        Private Const ELEMENT_Coeff As String = "Coeff"
        Private Const ELEMENT_IdxShapefile As String = "IdxShapefile"
        Private Const ELEMENT_IdxField As String = "IdxField"
        Private Const ELEMENT_Threshold As String = "Threshold"
        Private Const ELEMENT_TypeDefXmlFile As String = "TypeDefXmlFile"

        Public intID As Short
        Public intApply As Short
        Public strPollName As String
        Public strCoeffSet As String
        Public strCoeff As String
        Public idxShapefile As String
        Public idxField As String
        Public intThreshold As Short
        Public strTypeDefXmlFile As String

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
                NodeAppendChildElement(dom, node, ELEMENT_PollName, strPollName)
                NodeAppendChildElement(dom, node, ELEMENT_CoeffSet, strCoeffSet)
                NodeAppendChildElement(dom, node, ELEMENT_Coeff, strCoeff)
                NodeAppendChildElement(dom, node, ELEMENT_IdxShapefile, idxShapefile)
                NodeAppendChildElement(dom, node, ELEMENT_IdxField, idxField)
                NodeAppendChildElement(dom, node, ELEMENT_Threshold, intThreshold)
                NodeAppendChildElement(dom, node, ELEMENT_TypeDefXmlFile, strTypeDefXmlFile)
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

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
                strPollName = GetNodeText(node, ELEMENT_PollName)
                strCoeffSet = GetNodeText(node, ELEMENT_CoeffSet)
                strCoeff = GetNodeText(node, ELEMENT_Coeff)
                idxShapefile = GetNodeText(node, ELEMENT_IdxShapefile)
                idxField = GetNodeText(node, ELEMENT_IdxField)
                intThreshold = CShort(GetNodeText(node, ELEMENT_Threshold))
                strTypeDefXmlFile = GetNodeText(node, ELEMENT_TypeDefXmlFile)

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub
    End Class
End Namespace