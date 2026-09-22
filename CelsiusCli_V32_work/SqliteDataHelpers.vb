Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.Globalization

Module SqliteDataHelpers
    Public Function CreateParam(name As String, value As Object) As SQLiteParameter
        Return New SQLiteParameter(name, If(value Is Nothing, DBNull.Value, value))
    End Function

    Public Function LoadTable(connection As SQLiteConnection, sql As String, ParamArray parameters() As SQLiteParameter) As DataTable
        Using command As New SQLiteCommand(sql, connection)
            For Each parameter As SQLiteParameter In parameters
                command.Parameters.Add(parameter)
            Next

            Using adapter As New SQLiteDataAdapter(command)
                Dim table As New DataTable()
                adapter.Fill(table)
                Return table
            End Using
        End Using
    End Function

    Public Function LoadSingleRow(connection As SQLiteConnection, sql As String, ParamArray parameters() As SQLiteParameter) As DataRow
        Dim table = LoadTable(connection, sql, parameters)
        If table.Rows.Count = 0 Then
            Return Nothing
        End If

        Return table.Rows(0)
    End Function

    Public Function DbString(value As Object) As String
        If value Is Nothing OrElse value Is DBNull.Value Then
            Return String.Empty
        End If

        Return Convert.ToString(value, CultureInfo.InvariantCulture)
    End Function

    Public Function DbInt(value As Object) As Integer
        If value Is Nothing OrElse value Is DBNull.Value OrElse DbString(value) = String.Empty Then
            Return 0
        End If

        Return Convert.ToInt32(value, CultureInfo.InvariantCulture)
    End Function

    Public Function DbLong(value As Object) As Long
        If value Is Nothing OrElse value Is DBNull.Value OrElse DbString(value) = String.Empty Then
            Return 0
        End If

        Return Convert.ToInt64(value, CultureInfo.InvariantCulture)
    End Function

    Public Function DbDouble(value As Object) As Double
        If value Is Nothing OrElse value Is DBNull.Value OrElse DbString(value) = String.Empty Then
            Return 0
        End If

        Return Convert.ToDouble(value, CultureInfo.InvariantCulture)
    End Function

    Public Function DbBool(value As Object) As Boolean
        If value Is Nothing OrElse value Is DBNull.Value OrElse DbString(value) = String.Empty Then
            Return False
        End If

        If TypeOf value Is Boolean Then
            Return CBool(value)
        End If

        Dim text = DbString(value)
        If String.Equals(text, "true", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If String.Equals(text, "false", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        Return Convert.ToDouble(value, CultureInfo.InvariantCulture) <> 0
    End Function
End Module
