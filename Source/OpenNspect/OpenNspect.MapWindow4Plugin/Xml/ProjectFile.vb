'********************************************************************************************************
'File Name: XmlPrjFile.vb
'Description: Class for handling the project file xml
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
Imports System.Collections.Generic
Imports System.Xml
Imports System.IO

Namespace Xml

    Friend Class ProjectFile
        Inherits Base
        ' *************************************************************************************
        ' *  Perot Systems Government Services
        ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
        ' *  WrapperMain
        ' *************************************************************************************
        ' *  Description: Xml Wrapper for use with main form's variables
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

        'Following are the names of the NODES

        Private Const NODE_NAME As String = "NSPECTProjectFile"
        Private Const NODE_PRJNAME As String = "PrjName"
        Private Const NODE_PRJROOT As String = "PrjRoot"
        Private Const NODE_PRJWORKSPACE As String = "PrjWorkspace"
        Private Const NODE_LCGridName As String = "LCGridName"
        Private Const NODE_LCGridFileName As String = "LCGridFileName"
        Private Const NODE_LCGridUnits As String = "LCGridUnits"
        Private Const NODE_LCGridType As String = "LCGridType"
        Private Const NODE_SoilsDefName As String = "SoilsDefName"
        Private Const NODE_SoilsHydFileName As String = "SoilsHydFileName"
        Private Const NODE_SoilsKFileName As String = "SoilsKFileName"
        Private Const NODE_PrecipScenario As String = "PrecipScenario"
        Private Const NODE_WaterShedDelin As String = "WaterShedDelineation"
        Private Const NODE_WaterQuality As String = "WaterQualityStandard"
        Private Const NODE_SelectedPolys As String = "SelectedPolygons"
        Private Const NODE_SelectedPolyLyrName As String = "SelectedPolyLyrName"
        Private Const NODE_SelectedPolyFileName As String = "SelectedPolyFileName"
        Private Const NODE_SelectedPolyList As String = "SelectedPolyList"
        Private Const NODE_LocalEffects As String = "LocalEffects"
        Private Const NODE_CalcErosion As String = "CalcErosion"
        Private Const NODE_UseOWNSDR As String = "SDRGrid"
        Private Const NODE_SDRGridFileName As String = "SDRGridFileName"
        Private Const NODE_RainGridBool As String = "RainGridBool"
        Private Const NODE_RainGridName As String = "RainGridName"
        Private Const NODE_RainGridFileName As String = "RainGridFileName"
        Private Const NODE_RainConstBool As String = "RainConstBool"
        Private Const NODE_RainConstValue As String = "RainConstValue"

        Private _projectFilePath As String
        'Variables holding value of nodes above
        Private _ProjectName As String
        Private _ProjectRoot As String
        Private _ProjectWorkspace As String
        Private _LandCoverGridName As String
        Private _LandCoverGridDirectory As String
        Private _LandCoverGridUnits As String
        Private _LandCoverGridType As String
        Private _SoilsDefName As String
        Private _SoilsHydDirectory As String
        Private _SoilsKFileName As String
        Private _PrecipScenario As String
        Private _WaterShedDelin As String
        Private _WaterQuality As String
        Private _intSelectedPolys As Short
        Private _strSelectedPolyFileName As String
        Private _strSelectedPolyLyrName As String
        Public intSelectedPolyList As New List(Of Integer)
        Private _intLocalEffects As Short

        Private _intCalcErosion As Short
        Private _intUseOwnSDR As Short
        Private _strSDRGridFileName As String
        Private _intRainGridBool As Short
        Private _strRainGridName As String
        Private _strRainGridFileName As String
        Private _intRainConstBool As Short
        Private _dblRainConstValue As Double

        Public MgmtScenHolder As ManagementScenarioItems
        Public PollItems As PollutantItems
        Public LUItems As LandUseItems
        Public OutputItems As OutputItems

        Public Property ProjectName As String
            Get
                Return _ProjectName
            End Get
            Set(value As String)
                If _ProjectName = value Then
                    Return
                End If
                _ProjectName = value
            End Set
        End Property
        Public Property ProjectRoot As String
            Get
                ' assign default value
                If String.IsNullOrEmpty(_ProjectRoot) Then
                    Return "C:\NSPECT"
                Else
                    Return _ProjectRoot
                End If
            End Get
            Set(value As String)
                If _ProjectRoot = value Then
                    Return
                End If
                _ProjectRoot = value
            End Set
        End Property
        Public Property ProjectWorkspace As String
            Get
                Return _ProjectWorkspace
            End Get
            Set(value As String)
                If _ProjectWorkspace = value Then
                    Return
                End If
                _ProjectWorkspace = value
            End Set
        End Property
        Public Property LandCoverGridName As String
            Get
                Return _LandCoverGridName
            End Get
            Set(value As String)
                If _LandCoverGridName = value Then
                    Return
                End If
                _LandCoverGridName = value
            End Set
        End Property
        ''' <summary>
        ''' Tries to find a valid, relative path.
        ''' </summary>
        ''' <param name="value">The value.</param><returns></returns>
        Private Function EncourageValidPath(ByVal value As String) As String
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentException("value is nothing or empty.", "value")
            End If

            If Not Directory.Exists(value) Then
                Dim relativePath = value.Replace(ProjectRoot, String.Empty)
                Dim suggestedPath = Path.GetDirectoryName(_projectFilePath) + "\.." + relativePath
                If Directory.Exists(suggestedPath) Then
                    Return suggestedPath
                End If
            End If
            Return value
        End Function
        Public Property LandCoverGridDirectory As String
            Get
                Return _LandCoverGridDirectory
            End Get
            Set(value As String)
                If _LandCoverGridDirectory = value Then
                    Return
                End If

                _LandCoverGridDirectory = EncourageValidPath(value)

            End Set
        End Property
        Public Property LandCoverGridUnits As String
            Get
                Return _LandCoverGridUnits
            End Get
            Set(value As String)
                If _LandCoverGridUnits = value Then
                    Return
                End If
                _LandCoverGridUnits = value
            End Set
        End Property
        Public Property LandCoverGridType As String
            Get
                Return _LandCoverGridType
            End Get
            Set(value As String)
                If _LandCoverGridType = value Then
                    Return
                End If
                _LandCoverGridType = value
            End Set
        End Property
        Public Property SoilsDefName As String
            Get
                Return _SoilsDefName
            End Get
            Set(value As String)
                If _SoilsDefName = value Then
                    Return
                End If
                _SoilsDefName = value
            End Set
        End Property
        Public Property SoilsHydDirectory As String
            Get
                Return _SoilsHydDirectory
            End Get
            Set(value As String)
                If _SoilsHydDirectory = value Then
                    Return
                End If
                _SoilsHydDirectory = EncourageValidPath(value)
            End Set
        End Property
        Public Property SoilsKFileName As String
            Get
                Return _SoilsKFileName
            End Get
            Set(value As String)
                If _SoilsKFileName = value Then
                    Return
                End If
                _SoilsKFileName = value
            End Set
        End Property
        Public Property PrecipScenario As String
            Get
                Return _PrecipScenario
            End Get
            Set(value As String)
                If _PrecipScenario = value Then
                    Return
                End If
                _PrecipScenario = value
            End Set
        End Property
        Public Property WaterShedDelin As String
            Get
                Return _WaterShedDelin
            End Get
            Set(value As String)
                If _WaterShedDelin = value Then
                    Return
                End If
                _WaterShedDelin = value
            End Set
        End Property
        Public Property WaterQuality As String
            Get
                Return _WaterQuality
            End Get
            Set(value As String)
                If _WaterQuality = value Then
                    Return
                End If
                _WaterQuality = value
            End Set
        End Property
        Public Property IntSelectedPolys As Short
            Get
                Return _intSelectedPolys
            End Get
            Set(value As Short)
                If _intSelectedPolys = value Then
                    Return
                End If
                _intSelectedPolys = value
            End Set
        End Property
        Public ReadOnly Property UseSelectedPolygons As Boolean
            Get
                Return IntSelectedPolys = 1
            End Get
        End Property

        Public Property StrSelectedPolyFileName As String
            Get
                Return _strSelectedPolyFileName
            End Get
            Set(value As String)
                If _strSelectedPolyFileName = value Then
                    Return
                End If
                _strSelectedPolyFileName = value
            End Set
        End Property
        Public Property StrSelectedPolyLyrName As String
            Get
                Return _strSelectedPolyLyrName
            End Get
            Set(value As String)
                If _strSelectedPolyLyrName = value Then
                    Return
                End If
                _strSelectedPolyLyrName = value
            End Set
        End Property
        Public Property IntLocalEffects As Short
            Get
                Return _intLocalEffects
            End Get
            Set(value As Short)
                If _intLocalEffects = value Then
                    Return
                End If
                _intLocalEffects = value
            End Set
        End Property

        Public ReadOnly Property IncludeLocalEffects As Boolean
            Get
                Return IntLocalEffects = 1
            End Get
        End Property

        Public Property IntCalcErosion As Short
            Get
                Return _intCalcErosion
            End Get
            Set(value As Short)
                If _intCalcErosion = value Then
                    Return
                End If
                _intCalcErosion = value
            End Set
        End Property
        Public Property IntUseOwnSDR As Short
            Get
                Return _intUseOwnSDR
            End Get
            Set(value As Short)
                If _intUseOwnSDR = value Then
                    Return
                End If
                _intUseOwnSDR = value
            End Set
        End Property
        Public Property StrSDRGridFileName As String
            Get
                Return _strSDRGridFileName
            End Get
            Set(value As String)
                If _strSDRGridFileName = value Then
                    Return
                End If
                _strSDRGridFileName = value
            End Set
        End Property
        Public Property IntRainGridBool As Short
            Get
                Return _intRainGridBool
            End Get
            Set(value As Short)
                If _intRainGridBool = value Then
                    Return
                End If
                _intRainGridBool = value
            End Set
        End Property
        Public Property StrRainGridName As String
            Get
                Return _strRainGridName
            End Get
            Set(value As String)
                If _strRainGridName = value Then
                    Return
                End If
                _strRainGridName = value
            End Set
        End Property
        Public Property StrRainGridFileName As String
            Get
                Return _strRainGridFileName
            End Get
            Set(value As String)
                If _strRainGridFileName = value Then
                    Return
                End If
                _strRainGridFileName = value
            End Set
        End Property
        Public Property IntRainConstBool As Short
            Get
                Return _intRainConstBool
            End Get
            Set(value As Short)
                If _intRainConstBool = value Then
                    Return
                End If
                _intRainConstBool = value
            End Set
        End Property
        Public Property DblRainConstValue As Double
            Get
                Return _dblRainConstValue
            End Get
            Set(value As Double)
                If _dblRainConstValue = value Then
                    Return
                End If
                _dblRainConstValue = value
            End Set
        End Property


        Public ReadOnly Property NodeName() As String
            Get
                'Retrieve the name of the element that this class wraps.

                NodeName = NODE_NAME

            End Get
        End Property

        Public Overrides Property Xml() As String
            Get
                'Retrieve the Xml string that this class represents. The Xml returned is
                'built from the values of this class's properties.

                Xml = Me.CreateNode().OuterXml

            End Get
            Set(ByVal Value As String)
                'Assign a new Xml string to this class. The newly assigned Xml is parsed,
                'and the class's properties are set accordingly.

                Dim dom As New XmlDocument
                Dim node As XmlNode

                If InStr(Value, ".xml") > 0 Then
                    _projectFilePath = Value
                    dom.Load(Value)
                Else
                    dom.LoadXml(Value)
                End If

                node = dom.DocumentElement

                LoadNode(node)

            End Set
        End Property

        Public Sub SaveFile(ByRef strXml As String)
            Try
                Dim dom As New XmlDocument
                dom.LoadXml(Me.Xml)

                dom.Save(strXml)

            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

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
                    dom.LoadXml(String.Format("<{0}/>", NODE_NAME))
                    node = dom.DocumentElement
                    'Otherwise use passed-in parent.
                Else
                    dom = Parent.OwnerDocument
                    node = dom.CreateElement(NODE_NAME)
                    Parent.AppendChild(node)
                End If

                '*********************************************************************
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
                NodeAppendChildElement(dom, node, NODE_PRJNAME, ProjectName)
                NodeAppendChildElement(dom, node, NODE_PRJROOT, ProjectRoot)
                NodeAppendChildElement(dom, node, NODE_PRJWORKSPACE, ProjectWorkspace)
                NodeAppendChildElement(dom, node, NODE_LCGridName, LandCoverGridName)
                NodeAppendChildElement(dom, node, NODE_LCGridFileName, LandCoverGridDirectory)
                NodeAppendChildElement(dom, node, NODE_LCGridUnits, LandCoverGridUnits)
                NodeAppendChildElement(dom, node, NODE_LCGridType, LandCoverGridType)
                NodeAppendChildElement(dom, node, NODE_SoilsDefName, SoilsDefName)
                NodeAppendChildElement(dom, node, NODE_SoilsHydFileName, SoilsHydDirectory)
                NodeAppendChildElement(dom, node, NODE_SoilsKFileName, SoilsKFileName)
                'NodeAppendChildElement dom, node, NODE_RainFallType, intRainFallType
                NodeAppendChildElement(dom, node, NODE_PrecipScenario, PrecipScenario)
                NodeAppendChildElement(dom, node, NODE_WaterShedDelin, WaterShedDelin)
                NodeAppendChildElement(dom, node, NODE_WaterQuality, WaterQuality)
                NodeAppendChildElement(dom, node, NODE_SelectedPolys, IntSelectedPolys)
                NodeAppendChildElement(dom, node, NODE_SelectedPolyFileName, StrSelectedPolyFileName)
                Dim strlist As String = ""
                If intSelectedPolyList.Count > 0 Then strlist = intSelectedPolyList(0).ToString
                For i As Integer = 1 To intSelectedPolyList.Count - 1
                    strlist = strlist + "," + intSelectedPolyList(i).ToString
                Next
                NodeAppendChildElement(dom, node, NODE_SelectedPolyList, strlist)
                NodeAppendChildElement(dom, node, NODE_SelectedPolyLyrName, StrSelectedPolyLyrName)
                NodeAppendChildElement(dom, node, NODE_LocalEffects, IntLocalEffects)
                NodeAppendChildElement(dom, node, NODE_CalcErosion, IntCalcErosion)
                NodeAppendChildElement(dom, node, NODE_UseOWNSDR, IntUseOwnSDR)
                NodeAppendChildElement(dom, node, NODE_SDRGridFileName, StrSDRGridFileName)
                NodeAppendChildElement(dom, node, NODE_RainGridBool, IntRainGridBool)
                NodeAppendChildElement(dom, node, NODE_RainGridName, StrRainGridName)
                NodeAppendChildElement(dom, node, NODE_RainGridFileName, StrRainGridFileName)
                NodeAppendChildElement(dom, node, NODE_RainConstBool, IntRainConstBool)
                NodeAppendChildElement(dom, node, NODE_RainConstValue, DblRainConstValue)

                'Format
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

                'Pollutants
                PollItems.CreateNode(node)

                'Format
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

                'Management Scenarios
                MgmtScenHolder.CreateNode(node)

                'Format
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
                LUItems.CreateNode(node)

                'Format
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
                OutputItems.CreateNode(node)

                'Format
                node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

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

                ProjectName = GetNodeText(node, NODE_PRJNAME)
                ProjectRoot = GetNodeText(node, NODE_PRJROOT)
                ProjectWorkspace = GetNodeText(node, NODE_PRJWORKSPACE)
                LandCoverGridName = GetNodeText(node, NODE_LCGridName)
                LandCoverGridDirectory = GetNodeText(node, NODE_LCGridFileName)
                LandCoverGridUnits = GetNodeText(node, NODE_LCGridUnits)
                LandCoverGridType = GetNodeText(node, NODE_LCGridType)
                SoilsDefName = GetNodeText(node, NODE_SoilsDefName)
                SoilsHydDirectory = GetNodeText(node, NODE_SoilsHydFileName)
                SoilsKFileName = GetNodeText(node, NODE_SoilsKFileName)
                PrecipScenario = GetNodeText(node, NODE_PrecipScenario)
                WaterShedDelin = GetNodeText(node, NODE_WaterShedDelin)
                WaterQuality = GetNodeText(node, NODE_WaterQuality)
                IntSelectedPolys = CShort(GetNodeText(node, NODE_SelectedPolys))
                StrSelectedPolyFileName = GetNodeText(node, NODE_SelectedPolyFileName)
                Dim tmpstr As String() = GetNodeText(node, NODE_SelectedPolyList).Split(",")
                intSelectedPolyList.Clear()
                For i As Integer = 0 To tmpstr.Length - 1
                    If tmpstr(i) <> "" Then
                        intSelectedPolyList.Add(CShort(tmpstr(i)))
                    End If
                Next
                IntLocalEffects = CShort(GetNodeText(node, NODE_LocalEffects))
                IntCalcErosion = CShort(GetNodeText(node, NODE_CalcErosion))
                IntUseOwnSDR = CShort(GetNodeText(node, NODE_UseOWNSDR, "integer"))
                StrSDRGridFileName = GetNodeText(node, NODE_SDRGridFileName)
                IntRainGridBool = CShort(GetNodeText(node, NODE_RainGridBool))
                StrRainGridName = GetNodeText(node, NODE_RainGridName)
                StrRainGridFileName = GetNodeText(node, NODE_RainGridFileName)
                IntRainConstBool = CShort(GetNodeText(node, NODE_RainConstBool))
                DblRainConstValue = CDbl(GetNodeText(node, NODE_RainConstValue))

                MgmtScenHolder.LoadNode(node.SelectSingleNode(MgmtScenHolder.NodeName))
                PollItems.LoadNode(node.SelectSingleNode(PollItems.NodeName))
                LUItems.LoadNode(node.SelectSingleNode(LUItems.NodeName))
                OutputItems.LoadNode(node.SelectSingleNode(OutputItems.NodeName))
            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub

        Public Sub New()
            Try
                MgmtScenHolder = New ManagementScenarioItems
                'A collection of management scenarios
                PollItems = New PollutantItems
                'A collection of Pollutants
                LUItems = New LandUseItems
                'A collection of landuses
                OutputItems = New OutputItems
                'A collection of outputs
            Catch ex As Exception
                HandleError(ex)
            End Try
        End Sub
    End Class
End Namespace