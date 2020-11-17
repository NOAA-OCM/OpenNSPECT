Module ConvertTo
    Public Function GetConverterToTauDemFromEsri() As RasterMathCellCalcNulls
        Return New RasterMathCellCalcNulls(AddressOf ConvertTo.TauDemFromEsri)
    End Function
    Public Function GetConverterToEsriFromTauDem() As RasterMathCellCalcNulls
        Return New RasterMathCellCalcNulls(AddressOf ConvertTo.EsriFromTauDem)
    End Function

    Private Function TauDemFromEsri(ByVal direction As Single, ByVal Input1Null As Single, ByVal Input2 As Single,
                                   ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single,
                                   ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single,
                                   ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Return TauDemFromEsri(direction, Input1Null)
    End Function

    Private Function EsriFromTauDem(ByVal direction As Single, ByVal Input1Null As Single, ByVal Input2 As Single,
                                         ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single,
                                         ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single,
                                         ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Return EsriFromTauDem(direction, Input1Null)
    End Function

    ''' <summary>
    ''' Converts the tau DEM to ESRI cell. ESRI is clockwise 1-128 from east. TAUDEM is 1-8 counter-clockwise from east
    ''' </summary>
    ''' <param name="direction">The direction.</param>
    ''' <param name="Input1Null"></param><returns></returns>
    Private Function EsriFromTauDem(ByVal direction As Single, ByVal Input1Null As Single) As Single
        If direction = 1 Then
            Return 1
        ElseIf direction = 8 Then
            Return 2
        ElseIf direction = 7 Then
            Return 4
        ElseIf direction = 6 Then
            Return 8
        ElseIf direction = 5 Then
            Return 16
        ElseIf direction = 4 Then
            Return 32
        ElseIf direction = 3 Then
            Return 64
        ElseIf direction = 2 Then
            Return 128
        ElseIf direction = Input1Null Then
            Return -1
        Else
            Return -1
        End If
    End Function
    Private Function TauDemFromEsri(ByVal direction As Single, ByVal Input1Null As Single) As Single
        'ESRI is clockwise 1-128 from east. TAUDEM is 1-8 counter-clockwise from east
        If direction = 1 Then
            Return 1
        ElseIf direction = 2 Then
            Return 8
        ElseIf direction = 4 Then
            Return 7
        ElseIf direction = 8 Then
            Return 6
        ElseIf direction = 16 Then
            Return 5
        ElseIf direction = 32 Then
            Return 4
        ElseIf direction = 64 Then
            Return 3
        ElseIf direction = 128 Then
            Return 2
        ElseIf direction = Input1Null Then
            Return -1
        Else
            Return -1
        End If
    End Function
End Module
