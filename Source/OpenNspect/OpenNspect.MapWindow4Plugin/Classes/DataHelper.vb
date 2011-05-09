Imports System.Data.OleDb

Public Class DataHelper
    Implements IDisposable

    Private _command As OleDbCommand
    Private _reader As OleDbDataReader
    Private _adapter As OleDbDataAdapter

    Public Function CloneCommand() As OleDbCommand
        Return _command.Clone()
    End Function

    Public Function GetCommand() As OleDbCommand
        Return _command
    End Function

    Public Function GetAdapter() As OleDbDataAdapter
        If _adapter Is Nothing Then
            _adapter = New OleDbDataAdapter(_command)
        End If

        Return _adapter
    End Function
    ''' <summary>
    ''' Returns the reader associated with this object or creates a new one if needed.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteReader() As OleDbDataReader
        If _reader Is Nothing Then
            _reader = _command.ExecuteReader()
        End If

        Return _reader
    End Function

    Sub ExecuteNonQuery()
        _command.ExecuteNonQuery()
    End Sub

    Public Sub New(query As String)
        _command = New OleDbCommand(query, g_DBConn)
    End Sub

#Region "IDisposable Support"

    Private disposedValue As Boolean

    ' To detect redundant calls

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If _reader IsNot Nothing Then
                    _reader.Dispose()
                End If

                If _adapter IsNot Nothing Then
                    _adapter.Dispose()
                End If

                If _command IsNot Nothing Then
                    _command.Dispose()
                End If
            End If

            _adapter = Nothing
            _reader = Nothing
            _command = Nothing
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region
End Class
