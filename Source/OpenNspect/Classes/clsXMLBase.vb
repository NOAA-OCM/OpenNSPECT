'********************************************************************************************************
'File Name: clsXMLBase.vb
'Description: A class for handling the NSPECT xml base structure
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

Public Class clsXMLBase
    Public Overridable Property XML() As String
        Get
            'Retrieve the XML string that this class represents. The XML returned is
            'built from the values of this class's properties.

            XML = Me.CreateNode().OuterXml

        End Get
        Set (ByVal Value As String)
            'Assign a new XML string to this class. The newly assigned XML is parsed,
            'and the class's properties are set accordingly.

            Dim dom As New XmlDocument
            dom.LoadXml (Value)
            LoadNode ((dom.DocumentElement))

        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Parent"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function CreateNode (Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Return Nothing
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Public Overridable Sub LoadNode(ByRef node As XmlNode)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="strNodeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NodeGetText (ByVal node As XmlNode, ByVal strNodeName As String) As String
        Try
            'The NodeGetText function uses selectSingleNode on the passed-in node to
            'find the child node having the given name. When found, the text of that
            'node is returned. If the child node is not found in the node, then an
            'empty string is returned.

            NodeGetText = node.SelectSingleNode (strNodeName).Value

        Catch ex As Exception
            HandleError (ex)
            NodeGetText = ""
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dom"></param>
    ''' <param name="node"></param>
    ''' <param name="strName"></param>
    ''' <param name="varValue"></param>
    ''' <remarks></remarks>
    Public Sub NodeAppendAttribute (ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal strName As String, _
                                    ByVal varValue As Object)
        Try
            'The NodeAppendAttribute subroutine creates an attribute having the given
            'name and value, and adds it to the given node's Attributes collection.

            Dim attr As XmlAttribute

            'Create a new attribute and set its value.
            attr = dom.CreateAttribute (strName)
            attr.Value = varValue

            node.Attributes.Append (attr)

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dom"></param>
    ''' <param name="node"></param>
    ''' <param name="Name"></param>
    ''' <param name="Value"></param>
    ''' <remarks></remarks>
    Public Sub NodeAppendChildElement (ByVal dom As XmlDocument, ByVal node As XmlNode, ByVal Name As String, _
                                       ByVal Value As Object)
        Try
            'The NodeAppendChildElement subroutine creates an element having the given
            'name and value, and adds the element as a child of the given node.

            Dim element As XmlElement

            'Create a new child element and set its value.
            element = dom.CreateElement (Name)
            element.InnerText = CStr (Value)

            'Append the new child element to the node.
            node.AppendChild (element)
            'Format nicely
            node.AppendChild (dom.CreateTextNode (vbNewLine & vbTab))

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Parent"></param>
    ''' <param name="ChildName"></param>
    ''' <param name="NodeTextType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNodeText (ByRef Parent As XmlNode, ByRef ChildName As String, _
                                 Optional ByRef NodeTextType As String = "") As String
        Try
            'The GetNodeText function retrieves the value of the given child element
            'within the given parent element. If the requested child element is not
            'found, then an empty string is returned.

            GetNodeText = Parent.SelectSingleNode (ChildName).InnerText
        Catch ex As Exception
            If NodeTextType = "integer" Then
                GetNodeText = "0"
            Else
                GetNodeText = ""
            End If

            HandleError (ex)
        End Try
    End Function
End Class
