Attribute VB_Name = "modLCTypeValues"
' *************************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modLCTypeValues
' *************************************************************************************************
' *  Description: Code mod for handling validation of landcover type
' *
' *         Function CheckGridValuesLCtype::  checks the cover factor grid values
' *         Function Rollback::  Delete's previous work if user so chooses.
' *
' *  Called By:  frmLCTypes, frmNewLCtypes
' **************************************************************************************************

Public Function CheckGridValuesLCType(aryValue() As Variant) As Boolean
    
    Dim i As Integer
    i = 0
    
        For i = 0 To UBound(aryValue)
            'Select Case i
            Select Case i
                
                Case 6
                    If aryValue(i) < 0 Or aryValue(i) > 1 Then
                        MsgBox "Cover Factor must be between 0 and 1.  Please check values", vbCritical, "Check Cover Number"
                        CheckGridValuesLCType = False
                    Else
                        CheckGridValuesLCType = True
                    End If
                        
          End Select
        
        Next i
         
End Function

Public Sub RollBack(strName As String)
    
    Dim strSQLSelect As String
    strSQLSelect = "SELECT * FROM LCTYPE where NAME LIKE '" & strName & "'"
    
    Dim rsRollback As ADODB.Recordset
    Set rsRollback = New ADODB.Recordset
    
    rsRollback.Open strSQLSelect, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    rsRollback.Delete adAffectCurrent
    
    'clean
    Set rsRollback = Nothing

End Sub
