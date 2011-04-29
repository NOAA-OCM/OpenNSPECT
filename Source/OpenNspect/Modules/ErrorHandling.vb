'********************************************************************************************************
'File Name: modErrorCodes.vb
'Description: Error strings and other common dialog strings
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


Module ErrorHandling
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modErrorCodes
    ' *************************************************************************************
    ' *  Description:  Definition of all hand made error handling
    ' *
    ' *
    ' *  Called By:  Various
    ' *************************************************************************************

    Public Const Err1 As String = "Warning: There is a problem with the value you just entered.  Please insure " & "your entry is numeric, in the range of 0 - 1, and contains no more than " & vbNewLine & "four (4) values after the decimal point.  "

    Public Const Err2 As String = "Values must be unique.  Please check the number entered in the following cell:  "

    Public Const Err3 As String = "Warning: There is a problem with the value you just entered.  Please insure " & "your entry is numeric, in the range of 0 - 1, and contains no more than " & vbNewLine & "three (3) values after the decimal point.  "

    Public Const Err4 As String = "The name you have chosen is already in use.  Please enter another."
    Public Const Err5 As String = "Warning: Values must be greater than or equal to 0."

    Public Const Err6 As String = "Warning: There is a problem with the value you just entered.  Please insure " & "your entry is numeric, in the range of 0 - 1000, and contains no more than " & vbNewLine & "four (4) values after the decimal point.  "

    Public intYesNo As Short
    Public Const strYesNo As String = "You have made changes to the data.  Would you like to save before continuing?"
    Public Const strYesNoTitle As String = "Save Changes?"

    Public Const strDefault As String = "Are you sure you want to delete the current settings and restore default CCAP data?"

    Public Const strDefaultTitle As String = "Restore Defaults Settings?"

    'Constants for File Dialogs
    Public Const MSG1TextFile As String = "<name> Text File(*.txt;*.csv)|*.txt;*.csv"
    Public Const MSG2 As String = "<name> File to Import"
    Public Const MSG3 As String = "Export <name> into Filename"
    Public Const MSG4AnalysisFile As String = "Analysis File (*.prj)"
    Public Const MSG5 As String = "Open Analysis File"
    Public Const MSG6ShapeFile As String = "ESRI Shapefile (*.shp)|*.shp"
    Public Const MSG7 As String = ""
    Public Const MSG8XmlFile As String = "Xml File(*.xml)|*.xml"

    Public Sub DisplayError(ByRef Error_Renamed As String, ByRef i As Short, ByRef j As Short)
        MsgBox(String.Format("{0}Row: {1}, Column: {2}", Error_Renamed, (i + 1), (j + 1)), MsgBoxStyle.Critical, "Warning")
    End Sub

    Public Sub HandleError(ByRef ex As Exception)
        Try
            Using errorBox As New ErrorForm(ex)
                errorBox.ShowDialog()
            End Using
            Trace.TraceError(ex.Message)
        Catch ex2 As Exception
            ' // couldn't show the dialog.
        End Try
    End Sub
End Module
