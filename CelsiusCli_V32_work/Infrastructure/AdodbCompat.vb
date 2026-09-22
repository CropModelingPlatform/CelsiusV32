Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Data.SQLite

Namespace ADODB
    Public Module Constants
        Public Const adOpenDynamic As Integer = 2
        Public Const adLockOptimistic As Integer = 3
    End Module

    Public Class Connection
        Private _sqlite As SQLiteConnection

        Public ReadOnly Property InnerConnection As SQLiteConnection
            Get
                Return _sqlite
            End Get
        End Property

        Public Sub Open(connectionString As String)
            _sqlite = New SQLiteConnection(connectionString)
            _sqlite.Open()
        End Sub

        Public Sub Close()
            If _sqlite IsNot Nothing Then
                _sqlite.Close()
            End If
        End Sub
    End Class

    Public Class Recordset
        Private _table As DataTable
        Private _pendingRow As DataRow
        Private _currentIndex As Integer
        Private _tableName As String
        Private _fields As RecordsetFields
        Private _connection As SQLiteConnection

        Public Sub New()
            _fields = New RecordsetFields(Me)
            _currentIndex = -1
        End Sub

        Public ReadOnly Property Fields As RecordsetFields
            Get
                Return _fields
            End Get
        End Property

        Default Public ReadOnly Property Item(columnName As String) As Object
            Get
                Return CurrentRowValue(columnName)
            End Get
        End Property

        Public ReadOnly Property EOF As Boolean
            Get
                Return _table Is Nothing OrElse _table.Rows.Count = 0 OrElse _currentIndex < 0 OrElse _currentIndex >= _table.Rows.Count
            End Get
        End Property

        Public Sub Open(source As String, connection As Connection, Optional cursorType As Object = Nothing, Optional lockType As Object = Nothing)
            Dim sql = NormalizeSource(source)
            _tableName = DetectTableName(source, sql)
            _connection = connection.InnerConnection
            Dim adapter = New SQLiteDataAdapter(sql, connection.InnerConnection)
            _table = New DataTable(_tableName)
            adapter.Fill(_table)
            _pendingRow = Nothing
            _currentIndex = If(_table.Rows.Count > 0, 0, -1)
        End Sub

        Public Sub Close()
            _table = Nothing
            _pendingRow = Nothing
            _currentIndex = -1
            _connection = Nothing
        End Sub

        Public Sub MoveFirst()
            _currentIndex = If(_table IsNot Nothing AndAlso _table.Rows.Count > 0, 0, -1)
        End Sub

        Public Sub MoveNext()
            If _table Is Nothing Then
                _currentIndex = -1
                Return
            End If

            If _currentIndex < _table.Rows.Count Then
                _currentIndex += 1
            End If
        End Sub

        Public Sub MovePrevious()
            If _table Is Nothing Then
                _currentIndex = -1
                Return
            End If

            _currentIndex -= 1
            If _currentIndex < 0 Then
                _currentIndex = -1
            End If
        End Sub

        Public Sub MoveLast()
            _currentIndex = If(_table IsNot Nothing AndAlso _table.Rows.Count > 0, _table.Rows.Count - 1, -1)
        End Sub

        Public Sub Delete()
            If EOF Then
                Return
            End If

            ExecuteNonQuery($"DELETE FROM [{_tableName}]")
            _table.Rows.Clear()
            _currentIndex = -1
        End Sub

        Public Sub AddNew()
            EnsureTableLoaded()
            _pendingRow = _table.NewRow()
            _currentIndex = _table.Rows.Count
        End Sub

        Public Sub Update()
            EnsureTableLoaded()

            If _pendingRow IsNot Nothing Then
                _table.Rows.Add(_pendingRow)
                InsertPendingRow(_pendingRow)
                _pendingRow = Nothing
            End If

            If _currentIndex >= _table.Rows.Count Then
                _currentIndex = _table.Rows.Count - 1
            End If
        End Sub

        Friend Function GetField(index As Integer) As Field
            EnsureTableLoaded()
            Return New Field(Me, _table.Columns(index).ColumnName)
        End Function

        Friend Function GetField(columnName As String) As Field
            EnsureTableLoaded()
            Return New Field(Me, ResolveColumnName(columnName))
        End Function

        Friend Function CurrentRowValue(columnName As String) As Object
            Dim row = GetEditableRow()
            Dim value = row(ResolveColumnName(columnName))
            Return If(value Is DBNull.Value, Nothing, value)
        End Function

        Friend Sub SetCurrentRowValue(columnName As String, value As Object)
            Dim row = GetEditableRow()
            row(ResolveColumnName(columnName)) = If(value Is Nothing, DBNull.Value, value)
        End Sub

        Private Function GetEditableRow() As DataRow
            EnsureTableLoaded()

            If _pendingRow IsNot Nothing Then
                Return _pendingRow
            End If

            If EOF Then
                Throw New InvalidOperationException("Recordset cursor is not on a valid row.")
            End If

            Return _table.Rows(_currentIndex)
        End Function

        Private Sub EnsureTableLoaded()
            If _table Is Nothing Then
                Throw New InvalidOperationException("Recordset is not open.")
            End If
        End Sub

        Private Shared Function NormalizeSource(source As String) As String
            Dim trimmed = source.Trim()
            If Regex.IsMatch(trimmed, "^(SELECT|WITH)\b", RegexOptions.IgnoreCase) Then
                Return trimmed
            End If

            Return $"SELECT * FROM [{trimmed}]"
        End Function

        Private Shared Function DetectTableName(source As String, sql As String) As String
            Dim trimmed = source.Trim()
            If Not Regex.IsMatch(trimmed, "^(SELECT|WITH)\b", RegexOptions.IgnoreCase) Then
                Return trimmed
            End If

            Dim match = Regex.Match(sql, "FROM\s+\[?(?<name>[A-Za-z0-9_]+)\]?", RegexOptions.IgnoreCase)
            If match.Success Then
                Return match.Groups("name").Value
            End If

            Return "Result"
        End Function

        Private Function ResolveColumnName(columnName As String) As String
            For Each column As DataColumn In _table.Columns
                If String.Equals(column.ColumnName, columnName, StringComparison.OrdinalIgnoreCase) Then
                    Return column.ColumnName
                End If
            Next

            Throw New IndexOutOfRangeException($"Column '{columnName}' was not found in recordset '{_tableName}'.")
        End Function

        Private Sub InsertPendingRow(row As DataRow)
            Dim columns = New List(Of String)()
            Dim parameters = New List(Of String)()
            Dim command = New SQLiteCommand()
            command.Connection = _connection

            For index = 0 To _table.Columns.Count - 1
                Dim column = _table.Columns(index)
                Dim parameterName = $"@p{index}"
                columns.Add($"[{column.ColumnName}]")
                parameters.Add(parameterName)
                command.Parameters.AddWithValue(parameterName, row(column))
            Next

            command.CommandText = $"INSERT INTO [{_tableName}] ({String.Join(", ", columns)}) VALUES ({String.Join(", ", parameters)})"
            command.ExecuteNonQuery()
        End Sub

        Private Sub ExecuteNonQuery(sql As String)
            Using command = New SQLiteCommand(sql, _connection)
                command.ExecuteNonQuery()
            End Using
        End Sub
    End Class

    Public Class RecordsetFields
        Private ReadOnly _recordset As Recordset

        Public Sub New(recordset As Recordset)
            _recordset = recordset
        End Sub

        Default Public ReadOnly Property Item(index As Integer) As Field
            Get
                Return _recordset.GetField(index)
            End Get
        End Property

        Default Public ReadOnly Property Item(columnName As String) As Field
            Get
                Return _recordset.GetField(columnName)
            End Get
        End Property
    End Class

    Public Class Field
        Private ReadOnly _recordset As Recordset
        Private ReadOnly _columnName As String

        Public Sub New(recordset As Recordset, columnName As String)
            _recordset = recordset
            _columnName = columnName
        End Sub

        Public Property Value As Object
            Get
                Return _recordset.CurrentRowValue(_columnName)
            End Get
            Set(value As Object)
                _recordset.SetCurrentRowValue(_columnName, value)
            End Set
        End Property
    End Class
End Namespace
