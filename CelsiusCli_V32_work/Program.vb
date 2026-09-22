Imports System

Module Program
    Function Main(args As String()) As Integer
        If args.Length = 0 Then
            PrintUsage()
            Return 1
        End If

        Select Case args(0).ToLowerInvariant()
            Case "-h", "--help", "/?", "help"
                PrintUsage()
                Return 0
        End Select

        Try
            CelsiusRuntime.Run(args(0))
            Return 0
        Catch ex As Exception
            Console.Error.WriteLine(ex.ToString())
            Return 2
        End Try
    End Function

    Private Sub PrintUsage()
        Console.WriteLine("CELSIUS command-line runner")
        Console.WriteLine()
        Console.WriteLine("Usage:")
        Console.WriteLine("  celsiusV32 <sqlite-db-path>")
        Console.WriteLine("  celsiusV32 --help")
        Console.WriteLine()
        Console.WriteLine("Arguments:")
        Console.WriteLine("  <sqlite-db-path>  Path to the SQLite input database.")
        Console.WriteLine()
        Console.WriteLine("Behavior:")
        Console.WriteLine("  - Reads CELSIUS input tables from the SQLite database.")
        Console.WriteLine("  - Creates or reuses OutputSynt, OutputD_1 and OutputD_2.")
        Console.WriteLine("  - Runs all simulations listed in SimUnitList.")
    End Sub
End Module
