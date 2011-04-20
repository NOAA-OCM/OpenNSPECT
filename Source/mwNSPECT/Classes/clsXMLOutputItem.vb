'********************************************************************************************************
'File Name: clsXMLOutputItem.vb
'Description: Class for handling output item xml
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

Public Class clsXMLOutputItem
    Inherits clsXMLBase

    Const c_sModuleFileName As String = "clsXMLOutputItem.vb"

    Private Const NODE_NAME As String = "OutputFile"
    Private Const ELEMENT_OutputPath As String = "OutputPath"
    Private Const ELEMENT_OutputName As String = "OutputName"
    Private Const ELEMENT_OutputType As String = "OutputType"
    Private Const ELEMENT_OutputColor As String = "OutputColor"
    Private Const ELEMENT_OutputUseStretch As String = "OutputUseStretch"

    Public strPath As String
    Public strName As String
    Public strType As String
    Public strColor As String
    Public booUseStretch As Boolean

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

            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
            NodeAppendChildElement(dom, node, ELEMENT_OutputPath, strPath)
            NodeAppendChildElement(dom, node, ELEMENT_OutputName, strName)
            NodeAppendChildElement(dom, node, ELEMENT_OutputType, strType)
            NodeAppendChildElement(dom, node, ELEMENT_OutputColor, strColor)
            NodeAppendChildElement(dom, node, ELEMENT_OutputUseStretch, booUseStretch)

            'Return the created node
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

            strPath = GetNodeText(node, ELEMENT_OutputPath)
            strName = GetNodeText(node, ELEMENT_OutputName)
            strType = GetNodeText(node, ELEMENT_OutputType)
            strColor = GetNodeText(node, ELEMENT_OutputColor)
            If GetNodeText(node, ELEMENT_OutputUseStretch) <> "" Then
                booUseStretch = CType(GetNodeText(node, ELEMENT_OutputUseStretch), Boolean)
            Else
                booUseStretch = False
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

End Class