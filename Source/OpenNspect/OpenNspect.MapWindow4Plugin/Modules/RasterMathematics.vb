Imports System.IO
Imports MapWinGIS

Module RasterMathematics

    Public Delegate Function RasterMathCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single

    Public Delegate Function RasterMathCellCalcNulls(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single

    ''' <summary>
    ''' performs an operation using a number of lined up input grids on a cell-by-cell comparison.
    ''' </summary>
    ''' <param name="grid1">The grid1.</param>
    ''' <param name="grid2">The grid2.</param>
    ''' <param name="grid3">The grid3.</param>
    ''' <param name="grid4">The grid4.</param>
    ''' <param name="grid5">The grid5.</param>
    ''' <param name="outputGrid">The output grid.</param>
    ''' <param name="CellCalc">The cell calc.</param>
    ''' <param name="checkNullFirst">if set to <c>true</c> [check null first].</param>
    ''' <param name="CellCalcNull">The cell calc null.</param>
    Public Sub RasterMath(ByRef grid1 As Grid, ByRef grid2 As Grid, ByRef grid3 As Grid, ByRef grid4 As Grid, ByRef grid5 As Grid, ByRef outputGrid As Grid, ByRef CellCalc As RasterMathCellCalc, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcNull As RasterMathCellCalcNulls = Nothing)
        Const nodataout As Single = -9999.0
        Dim nodata1, nodata2, nodata3, nodata4, nodata5 As Single
        Dim tmppath As String = GetTempFileNameOutputGridExt()
        Dim grid1Header As GridHeader = grid1.Header

        Dim outputHeader As GridHeader = New GridHeader
        outputHeader.CopyFrom(grid1Header)
        outputHeader.NodataValue = nodataout

        outputGrid = New Grid()
        outputGrid.CreateNew(tmppath, outputHeader, GridDataType.FloatDataType, outputHeader.NodataValue)
        outputGrid.Header.Projection = grid1Header.Projection

        nodata1 = grid1Header.NodataValue
        nodata2 = If(grid2 Is Nothing, nodata1, grid2.Header.NodataValue)
        nodata3 = If(grid3 Is Nothing, nodata1, grid3.Header.NodataValue)
        nodata4 = If(grid4 Is Nothing, nodata1, grid4.Header.NodataValue)
        nodata5 = If(grid5 Is Nothing, nodata1, grid5.Header.NodataValue)

        Dim ncol As Integer = grid1Header.NumberCols - 1
        Dim rowvals1(ncol) As Single
        Dim rowvals2(ncol) As Single
        Dim rowvals3(ncol) As Single
        Dim rowvals4(ncol) As Single
        Dim rowvals5(ncol) As Single
        Dim rowvalsout(ncol) As Single

        Dim nrow As Integer = grid1Header.NumberRows - 1
        For row As Integer = 0 To nrow
            grid1.GetRow(row, rowvals1(0))

            If grid2 IsNot Nothing Then grid2.GetRow(row, rowvals2(0)) Else grid1.GetRow(row, rowvals2(0))
            If grid3 IsNot Nothing Then grid3.GetRow(row, rowvals3(0)) Else grid1.GetRow(row, rowvals3(0))
            If grid4 IsNot Nothing Then grid4.GetRow(row, rowvals4(0)) Else grid1.GetRow(row, rowvals4(0))
            If grid5 IsNot Nothing Then grid5.GetRow(row, rowvals5(0)) Else grid1.GetRow(row, rowvals5(0))

            For col As Integer = 0 To ncol
                If checkNullFirst Then
                    If rowvals1(col) = nodata1 OrElse rowvals2(col) = nodata2 OrElse rowvals3(col) = nodata3 OrElse rowvals4(col) = nodata4 OrElse rowvals5(col) = nodata5 Then
                        rowvalsout(col) = nodataout
                    Else
                        rowvalsout(col) = CellCalc.Invoke(rowvals1(col), rowvals2(col), rowvals3(col), rowvals4(col), rowvals5(col), nodataout)
                    End If
                Else
                    rowvalsout(col) = CellCalcNull.Invoke(rowvals1(col), nodata1, rowvals2(col), nodata2, rowvals3(col), nodata3, rowvals4(col), nodata4, rowvals5(col), nodata5, nodataout)
                End If
            Next

            outputGrid.PutRow(row, rowvalsout(0))
        Next
    End Sub


    Public Delegate Function RasterMathCellCalcWindow(ByRef InputBox1(,) As Single, ByRef InputBox2(,) As Single, ByRef InputBox3(,) As Single, ByRef InputBox4(,) As Single, ByRef InputBox5(,) As Single, ByVal OutNull As Single) As Single

    Public Delegate Function RasterMathCellCalcWindowNulls(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single

    ''' <summary>
    ''' performs an operation using a number of lined up input grids on a cell-by-cell comparison exposing the 8 cell surround the current center cell:
    ''' 7 0 1
    ''' 6 C 2
    ''' 5 4 3
    ''' </summary>
    ''' <param name="InputGrid1">The input grid1.</param>
    ''' <param name="InputGrid2">The input grid2.</param>
    ''' <param name="Inputgrid3">The inputgrid3.</param>
    ''' <param name="Inputgrid4">The inputgrid4.</param>
    ''' <param name="Inputgrid5">The inputgrid5.</param>
    ''' <param name="Outputgrid">The outputgrid.</param>
    ''' <param name="CellCalcWindow">The cell calc window.</param>
    ''' <param name="checkNullFirst">if set to <c>true</c> [check null first].</param>
    ''' <param name="CellCalcWindowNull">The cell calc window null.</param>
    Public Sub RasterMathWindow(ByRef InputGrid1 As Grid, ByRef InputGrid2 As Grid, ByRef Inputgrid3 As Grid, ByRef Inputgrid4 As Grid, ByRef Inputgrid5 As Grid, ByRef Outputgrid As Grid, ByRef CellCalcWindow As RasterMathCellCalcWindow, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcWindowNull As RasterMathCellCalcWindowNulls = Nothing)
        Dim head1, head2, head3, head4, head5, headnew As GridHeader
        Dim ncol As Integer
        Dim nrow As Integer
        Dim nodata1, nodata2, nodata3, nodata4, nodata5, nodataout As Single
        Dim rowvals1(3)(), rowvals2(3)(), rowvals3(3)(), rowvals4(3)(), rowvals5(3)(), rowvalsout() As Single
        Dim InputBox1(3, 3), InputBox2(3, 3), InputBox3(3, 3), InputBox4(3, 3), InputBox5(3, 3) As Single
        Dim tmppath As String = Path.GetTempFileName
        g_TempFilesToDel.Add(tmppath)
        tmppath = tmppath + OutputGridExt
        g_TempFilesToDel.Add(tmppath)

        nodataout = -9999.0

        head1 = InputGrid1.Header
        headnew = New GridHeader
        headnew.CopyFrom(head1)
        headnew.NodataValue = nodataout
        Outputgrid = New Grid()
        Outputgrid.CreateNew(tmppath, headnew, GridDataType.FloatDataType, headnew.NodataValue)
        Outputgrid.Header.Projection = head1.Projection
        ncol = head1.NumberCols - 1
        nrow = head1.NumberRows - 1
        nodata1 = head1.NodataValue
        If Not InputGrid2 Is Nothing Then
            head2 = InputGrid2.Header
            nodata2 = head2.NodataValue
        Else
            nodata2 = nodata1
        End If
        If Not Inputgrid3 Is Nothing Then
            head3 = Inputgrid3.Header
            nodata3 = head3.NodataValue
        Else
            nodata3 = nodata1
        End If
        If Not Inputgrid4 Is Nothing Then
            head4 = Inputgrid4.Header
            nodata4 = head4.NodataValue
        Else
            nodata4 = nodata1
        End If
        If Not Inputgrid5 Is Nothing Then
            head5 = Inputgrid5.Header
            nodata5 = head5.NodataValue
        Else
            nodata5 = nodata1
        End If
        For i As Integer = 0 To 2
            ReDim rowvals1(i)(ncol)
            ReDim rowvals2(i)(ncol)
            ReDim rowvals3(i)(ncol)
            ReDim rowvals4(i)(ncol)
            ReDim rowvals5(i)(ncol)
        Next

        ReDim rowvalsout(ncol)

        For row As Integer = 0 To nrow
            InputGrid1.GetRow(row, rowvals1(1)(0))
            If Not InputGrid2 Is Nothing Then
                InputGrid2.GetRow(row, rowvals2(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals2(1)(0))
            End If
            If Not Inputgrid3 Is Nothing Then
                Inputgrid3.GetRow(row, rowvals3(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals3(1)(0))
            End If
            If Not Inputgrid4 Is Nothing Then
                Inputgrid4.GetRow(row, rowvals4(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals4(1)(0))
            End If
            If Not Inputgrid5 Is Nothing Then
                Inputgrid5.GetRow(row, rowvals5(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals5(1)(0))
            End If

            If row <> 0 Then
                InputGrid1.GetRow(row - 1, rowvals1(0)(0))
                If Not InputGrid2 Is Nothing Then
                    InputGrid2.GetRow(row - 1, rowvals2(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals2(0)(0))
                End If
                If Not Inputgrid3 Is Nothing Then
                    Inputgrid3.GetRow(row - 1, rowvals3(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals3(0)(0))
                End If
                If Not Inputgrid4 Is Nothing Then
                    Inputgrid4.GetRow(row - 1, rowvals4(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals4(0)(0))
                End If
                If Not Inputgrid5 Is Nothing Then
                    Inputgrid5.GetRow(row - 1, rowvals5(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals5(0)(0))
                End If
            Else
                ReDim rowvals1(0)(rowvals1(1).Length)
                ReDim rowvals2(0)(rowvals1(1).Length)
                ReDim rowvals3(0)(rowvals1(1).Length)
                ReDim rowvals4(0)(rowvals1(1).Length)
                ReDim rowvals5(0)(rowvals1(1).Length)
            End If

            If row <> nrow Then
                InputGrid1.GetRow(row + 1, rowvals1(2)(0))
                If Not InputGrid2 Is Nothing Then
                    InputGrid2.GetRow(row + 1, rowvals2(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals2(2)(0))
                End If
                If Not Inputgrid3 Is Nothing Then
                    Inputgrid3.GetRow(row + 1, rowvals3(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals3(2)(0))
                End If
                If Not Inputgrid4 Is Nothing Then
                    Inputgrid4.GetRow(row + 1, rowvals4(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals4(2)(0))
                End If
                If Not Inputgrid5 Is Nothing Then
                    Inputgrid5.GetRow(row + 1, rowvals5(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals5(2)(0))
                End If
            Else
                ReDim rowvals1(2)(rowvals1(1).Length)
                ReDim rowvals2(2)(rowvals1(1).Length)
                ReDim rowvals3(2)(rowvals1(1).Length)
                ReDim rowvals4(2)(rowvals1(1).Length)
                ReDim rowvals5(2)(rowvals1(1).Length)
            End If

            For col As Integer = 1 To ncol - 1
                If row <> 0 Then
                    InputBox1(0, 0) = rowvals1(0)(col - 1)
                    InputBox1(0, 1) = rowvals1(0)(col)
                    InputBox1(0, 2) = rowvals1(0)(col + 1)
                    InputBox2(0, 0) = rowvals2(0)(col - 1)
                    InputBox2(0, 1) = rowvals2(0)(col)
                    InputBox2(0, 2) = rowvals2(0)(col + 1)
                    InputBox3(0, 0) = rowvals3(0)(col - 1)
                    InputBox3(0, 1) = rowvals3(0)(col)
                    InputBox3(0, 2) = rowvals3(0)(col + 1)
                    InputBox4(0, 0) = rowvals4(0)(col - 1)
                    InputBox4(0, 1) = rowvals4(0)(col)
                    InputBox4(0, 2) = rowvals4(0)(col + 1)
                    InputBox5(0, 0) = rowvals5(0)(col - 1)
                    InputBox5(0, 1) = rowvals5(0)(col)
                    InputBox5(0, 2) = rowvals5(0)(col + 1)
                Else
                    InputBox1(0, 0) = nodata1
                    InputBox1(0, 1) = nodata1
                    InputBox1(0, 2) = nodata1
                    InputBox2(0, 0) = nodata2
                    InputBox2(0, 1) = nodata2
                    InputBox2(0, 2) = nodata2
                    InputBox3(0, 0) = nodata3
                    InputBox3(0, 1) = nodata3
                    InputBox3(0, 2) = nodata3
                    InputBox4(0, 0) = nodata4
                    InputBox4(0, 1) = nodata4
                    InputBox4(0, 2) = nodata4
                    InputBox5(0, 0) = nodata5
                    InputBox5(0, 1) = nodata5
                    InputBox5(0, 2) = nodata5
                End If

                InputBox1(1, 0) = rowvals1(1)(col - 1)
                InputBox1(1, 1) = rowvals1(1)(col)
                InputBox1(1, 2) = rowvals1(1)(col + 1)
                InputBox2(1, 0) = rowvals2(1)(col - 1)
                InputBox2(1, 1) = rowvals2(1)(col)
                InputBox2(1, 2) = rowvals2(1)(col + 1)
                InputBox3(1, 0) = rowvals3(1)(col - 1)
                InputBox3(1, 1) = rowvals3(1)(col)
                InputBox3(1, 2) = rowvals3(1)(col + 1)
                InputBox4(1, 0) = rowvals4(1)(col - 1)
                InputBox4(1, 1) = rowvals4(1)(col)
                InputBox4(1, 2) = rowvals4(1)(col + 1)
                InputBox5(1, 0) = rowvals5(1)(col - 1)
                InputBox5(1, 1) = rowvals5(1)(col)
                InputBox5(1, 2) = rowvals5(1)(col + 1)

                If row <> 0 Then
                    InputBox1(2, 0) = rowvals1(2)(col - 1)
                    InputBox1(2, 1) = rowvals1(2)(col)
                    InputBox1(2, 2) = rowvals1(2)(col + 1)
                    InputBox2(2, 0) = rowvals2(2)(col - 1)
                    InputBox2(2, 1) = rowvals2(2)(col)
                    InputBox2(2, 2) = rowvals2(2)(col + 1)
                    InputBox3(2, 0) = rowvals3(2)(col - 1)
                    InputBox3(2, 1) = rowvals3(2)(col)
                    InputBox3(2, 2) = rowvals3(2)(col + 1)
                    InputBox4(2, 0) = rowvals4(2)(col - 1)
                    InputBox4(2, 1) = rowvals4(2)(col)
                    InputBox4(2, 2) = rowvals4(2)(col + 1)
                    InputBox5(2, 0) = rowvals5(2)(col - 1)
                    InputBox5(2, 1) = rowvals5(2)(col)
                    InputBox5(2, 2) = rowvals5(2)(col + 1)
                Else
                    InputBox1(2, 0) = nodata1
                    InputBox1(2, 1) = nodata1
                    InputBox1(2, 2) = nodata1
                    InputBox2(2, 0) = nodata2
                    InputBox2(2, 1) = nodata2
                    InputBox2(2, 2) = nodata2
                    InputBox3(2, 0) = nodata3
                    InputBox3(2, 1) = nodata3
                    InputBox3(2, 2) = nodata3
                    InputBox4(2, 0) = nodata4
                    InputBox4(2, 1) = nodata4
                    InputBox4(2, 2) = nodata4
                    InputBox5(2, 0) = nodata5
                    InputBox5(2, 1) = nodata5
                    InputBox5(2, 2) = nodata5
                End If

                If checkNullFirst Then
                    If InputBox1(0, 0) = nodata1 OrElse InputBox1(0, 1) = nodata1 OrElse InputBox1(0, 2) = nodata1 OrElse InputBox2(0, 0) = nodata2 OrElse InputBox2(0, 1) = nodata2 OrElse InputBox2(0, 2) = nodata2 OrElse InputBox3(0, 0) = nodata3 OrElse InputBox3(0, 1) = nodata3 OrElse InputBox3(0, 2) = nodata3 OrElse InputBox4(0, 0) = nodata4 OrElse InputBox4(0, 1) = nodata4 OrElse InputBox4(0, 2) = nodata4 OrElse InputBox5(0, 0) = nodata5 OrElse InputBox5(0, 1) = nodata5 OrElse InputBox5(0, 2) = nodata5 OrElse InputBox1(1, 0) = nodata1 OrElse InputBox1(1, 1) = nodata1 OrElse InputBox1(1, 2) = nodata1 OrElse InputBox2(1, 0) = nodata2 OrElse InputBox2(1, 1) = nodata2 OrElse InputBox2(1, 2) = nodata2 OrElse InputBox3(1, 0) = nodata3 OrElse InputBox3(1, 1) = nodata3 OrElse InputBox3(1, 2) = nodata3 OrElse InputBox4(1, 0) = nodata4 OrElse InputBox4(1, 1) = nodata4 OrElse InputBox4(1, 2) = nodata4 OrElse InputBox5(1, 0) = nodata5 OrElse InputBox5(1, 1) = nodata5 OrElse InputBox5(1, 2) = nodata5 OrElse InputBox1(2, 0) = nodata1 OrElse InputBox1(2, 1) = nodata1 OrElse InputBox1(2, 2) = nodata1 OrElse InputBox2(2, 0) = nodata2 OrElse InputBox2(2, 1) = nodata2 OrElse InputBox2(2, 2) = nodata2 OrElse InputBox3(2, 0) = nodata3 OrElse InputBox3(2, 1) = nodata3 OrElse InputBox3(2, 2) = nodata3 OrElse InputBox4(2, 0) = nodata4 OrElse InputBox4(2, 1) = nodata4 OrElse InputBox4(2, 2) = nodata4 OrElse InputBox5(2, 0) = nodata5 OrElse InputBox5(2, 1) = nodata5 OrElse InputBox5(2, 2) = nodata5 Then
                        rowvalsout(col) = nodataout
                    Else
                        rowvalsout(col) = CellCalcWindow.Invoke(InputBox1, InputBox2, InputBox3, InputBox4, InputBox5, nodataout)
                    End If
                Else
                    rowvalsout(col) = CellCalcWindowNull.Invoke(InputBox1, nodata1, InputBox2, nodata2, InputBox3, nodata3, InputBox4, nodata4, InputBox5, nodata5, nodataout)
                End If
            Next

            Outputgrid.PutRow(row, rowvalsout(0))
        Next
    End Sub
End Module
