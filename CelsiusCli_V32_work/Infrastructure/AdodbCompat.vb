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
        Private _transaction As SQLiteTransaction

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
            If _transaction IsNot Nothing Then
                RollbackTrans()
            End If
            If _sqlite IsNot Nothing Then
                _sqlite.Close()
            End If
        End Sub

        ' ADO-style transaction control. System.Data.SQLite applies an open
        ' transaction to every command issued on the same connection.
        Public Sub BeginTrans()
            If _transaction Is Nothing Then
                _transaction = _sqlite.BeginTransaction()
            End If
        End Sub

        Public Sub CommitTrans()
            If _transaction IsNot Nothing Then
                _transaction.Commit()
                _transaction.Dispose()
                _transaction = Nothing
            End If
        End Sub

        Public Sub RollbackTrans()
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
                _transaction.Dispose()
                _transaction = Nothing
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
        Private _sql As String
        ' Recordsets opened on a bare table name load their rows only when a
        ' row is actually read. Until then _table holds the schema only, so
        ' appending to a large output table does not reload it.
        Private _rowsLoaded As Boolean
        Private _lazyHasRows As Boolean
        Private _lazyAtLast As Boolean
        Private _insertCommand As SQLiteCommand

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
                If _table IsNot Nothing AndAlso Not _rowsLoaded Then
                    Return Not _lazyHasRows
                End If
                Return _table Is Nothing OrElse _table.Rows.Count = 0 OrElse _currentIndex < 0 OrElse _currentIndex >= _table.Rows.Count
            End Get
        End Property

        Public Sub Open(source As String, connection As Connection, Optional cursorType As Object = Nothing, Optional lockType As Object = Nothing)
            Dim sql = NormalizeSource(source)
            _tableName = DetectTableName(source, sql)
            _connection = connection.InnerConnection
            _sql = sql
            _pendingRow = Nothing
            DisposeInsertCommand()

            If IsBareTableName(source) Then
                _table = New DataTable(_tableName)
                Using adapter = New SQLiteDataAdapter($"SELECT * FROM [{_tableName}] LIMIT 0", _connection)
                    adapter.Fill(_table)
                End Using
                Using command = New SQLiteCommand($"SELECT EXISTS (SELECT 1 FROM [{_tableName}])", _connection)
                    _lazyHasRows = Convert.ToInt64(command.ExecuteScalar()) <> 0
                End Using
                _rowsLoaded = False
                _lazyAtLast = False
                _currentIndex = -1
            Else
                LoadRows(False)
            End If
        End Sub

        Private Sub LoadRows(atLast As Boolean)
            Dim loaded = New DataTable(_tableName)
            Using adapter = New SQLiteDataAdapter(_sql, _connection)
                adapter.Fill(loaded)
            End Using
            _table = loaded
            _rowsLoaded = True
            If _table.Rows.Count = 0 Then
                _currentIndex = -1
            ElseIf atLast Then
                _currentIndex = _table.Rows.Count - 1
            Else
                _currentIndex = 0
            End If
        End Sub

        Private Sub EnsureRowsLoaded()
            EnsureTableLoaded()
            If Not _rowsLoaded Then
                LoadRows(_lazyAtLast)
            End If
        End Sub

        Public Sub Close()
            DisposeInsertCommand()
            _table = Nothing
            _pendingRow = Nothing
            _currentIndex = -1
            _connection = Nothing
            _rowsLoaded = False
        End Sub

        Public Sub MoveFirst()
            If _table IsNot Nothing AndAlso Not _rowsLoaded Then
                _lazyAtLast = False
                Return
            End If
            _currentIndex = If(_table IsNot Nothing AndAlso _table.Rows.Count > 0, 0, -1)
        End Sub

        Public Sub MoveNext()
            If _table Is Nothing Then
                _currentIndex = -1
                Return
            End If

            EnsureRowsLoaded()
            If _currentIndex < _table.Rows.Count Then
                _currentIndex += 1
            End If
        End Sub

        Public Sub MovePrevious()
            If _table Is Nothing Then
                _currentIndex = -1
                Return
            End If

            EnsureRowsLoaded()
            _currentIndex -= 1
            If _currentIndex < 0 Then
                _currentIndex = -1
            End If
        End Sub

        Public Sub MoveLast()
            If _table IsNot Nothing AndAlso Not _rowsLoaded Then
                _lazyAtLast = True
                Return
            End If
            _currentIndex = If(_table IsNot Nothing AndAlso _table.Rows.Count > 0, _table.Rows.Count - 1, -1)
        End Sub

        Public Sub Delete()
            If EOF Then
                Return
            End If

            ExecuteNonQuery($"DELETE FROM [{_tableName}]")
            _table.Rows.Clear()
            _rowsLoaded = True
            _lazyHasRows = False
            _currentIndex = -1
        End Sub

        Public Sub AddNew()
            EnsureTableLoaded()
            _pendingRow = _table.NewRow()
            If _rowsLoaded Then
                _currentIndex = _table.Rows.Count
            End If
        End Sub

        Public Sub Update()
            EnsureTableLoaded()

            If _pendingRow IsNot Nothing Then
                InsertPendingRow(_pendingRow)
                If _rowsLoaded Then
                    _table.Rows.Add(_pendingRow)
                Else
                    ' Like ADO, the new record becomes the current (last) one.
                    _lazyHasRows = True
                    _lazyAtLast = True
                End If
                _pendingRow = Nothing
                Return
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

            EnsureRowsLoaded()
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

        Private Shared Function IsBareTableName(source As String) As Boolean
            Return Not Regex.IsMatch(source.Trim(), "^(SELECT|WITH)\b", RegexOptions.IgnoreCase)
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
            ' The INSERT is built once per open recordset and reused.
            If _insertCommand Is Nothing Then
                Dim columns = New List(Of String)()
                Dim parameters = New List(Of String)()
                _insertCommand = New SQLiteCommand()
                _insertCommand.Connection = _connection

                For index = 0 To _table.Columns.Count - 1
                    Dim parameterName = $"@p{index}"
                    columns.Add($"[{_table.Columns(index).ColumnName}]")
                    parameters.Add(parameterName)
                    _insertCommand.Parameters.Add(New SQLiteParameter(parameterName))
                Next

                _insertCommand.CommandText = $"INSERT INTO [{_tableName}] ({String.Join(", ", columns)}) VALUES ({String.Join(", ", parameters)})"
            End If

            For index = 0 To _table.Columns.Count - 1
                _insertCommand.Parameters(index).Value = row(index)
            Next
            _insertCommand.ExecuteNonQuery()
        End Sub

        Private Sub DisposeInsertCommand()
            If _insertCommand IsNot Nothing Then
                _insertCommand.Dispose()
                _insertCommand = Nothing
            End If
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
