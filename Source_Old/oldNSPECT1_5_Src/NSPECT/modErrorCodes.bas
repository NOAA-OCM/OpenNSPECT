Attribute VB_Name = "modErrorCodes"
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

Option Explicit

Public Const Err1 As String = "Warning: There is a problem with the value you just entered.  Please insure " & _
                              "your entry is numeric, in the range of 0 - 1, and contains no more than " & vbNewLine & _
                              "four (4) values after the decimal point.  "
Public Const Err2 As String = "Values must be unique.  Please check the number entered in the following cell:  "
Public Const Err3 As String = "Warning: There is a problem with the value you just entered.  Please insure " & _
                              "your entry is numeric, in the range of 0 - 1, and contains no more than " & vbNewLine & _
                              "three (3) values after the decimal point.  "
Public Const Err4 As String = "The name you have chosen is already in use.  Please enter another."
Public Const Err5 As String = "Warning: Values must be greater than or equal to 0."
Public Const Err6 As String = "Warning: There is a problem with the value you just entered.  Please insure " & _
                              "your entry is numeric, in the range of 0 - 1000, and contains no more than " & vbNewLine & _
                              "four (4) values after the decimal point.  "

Public intYesNo As Integer
Public Const strYesNo As String = "You have made changes to the data.  Would you like to save before continuing?"
Public Const strYesNoTitle As String = "Save Changes?"
Public Const strDefault As String = "Are you sure you want to delete the current settings and restore default CCAP data?"
Public Const strDefaultTitle As String = "Restore Defaults Settings?"

'Constants for File Dialogs
Public Const MSG1 = "<name> Text File(*.txt;*.csv)|*.txt;*.csv"
Public Const MSG2 = "<name> File to Import"
Public Const MSG3 = "Export <name> into Filename"
Public Const MSG4 = "Analysis File (*.prj)"
Public Const MSG5 = "Open Analysis File"
Public Const MSG6 = "ESRI Shapefile (*.shp)|*.shp"
Public Const MSG7 = ""
Public Const MSG8 = "XML File(*.xml)|*.xml"

Public Sub ErrorGenerator(Error As String, i As Integer, j As Integer)

    MsgBox Error & "Row: " & i & ", Column: " & j, vbCritical, "Warning"
    
End Sub

