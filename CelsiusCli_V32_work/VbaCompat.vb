Imports System

Module VbaCompat
    Public Const adOpenDynamic As Integer = ADODB.adOpenDynamic
    Public Const adLockOptimistic As Integer = ADODB.adLockOptimistic
    Public Const sigma As Double = 0.000000004903
    Public Codesuite As Integer

    Public Function IsNull(value As Object) As Boolean
        Return value Is Nothing OrElse value Is DBNull.Value
    End Function

    Public Function Sin(value As Double) As Double
        Return Math.Sin(value)
    End Function

    Public Function Cos(value As Double) As Double
        Return Math.Cos(value)
    End Function

    Public Function Tan(value As Double) As Double
        Return Math.Tan(value)
    End Function

    Public Function Atn(value As Double) As Double
        Return Math.Atan(value)
    End Function

    Public Function Sqr(value As Double) As Double
        Return Math.Sqrt(value)
    End Function

    Public Function Exp(value As Double) As Double
        Return Math.Exp(value)
    End Function

    Public Function Log(value As Double) As Double
        Return Math.Log(value)
    End Function
End Module
