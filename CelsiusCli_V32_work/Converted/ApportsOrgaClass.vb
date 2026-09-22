Option Strict Off
Option Explicit Off
Imports System
Public Class ApportsOrgaClass
    Dim DataBase_Cnn As ADODB.Connection
    'classe décrivanht les apports Organiques (résidus, fumiers)
    'VX: variable explicative, VE: variable d 'état, VS: variable simulée,PX: paramètre
    Dim Trouve As Boolean
    Dim Akres As Double ' parametre de la vitesse de décomposition des résidus
    Dim Bkres As Double ' parametre de la vitesse de décomposition des résidus
    Dim AHres As Double ' parametre de la vitesse de décomposition de la biomasse microbienne
    Dim BHres As Double ' parametre de la vitesse de décomposition de la biomasse microbienne
    Dim Nrec As Double 'teneur en N des résidus
    Dim AWB As Double 'parametre du C/N de la biomasse microbienne
    Dim BWB As Double 'parametre du C/N de la biomasse microbienne
    Dim Yres As Double 'constante de partition de la décomposition des apports organiques veres la biomasse microbienne
    Dim Kbio As Double ' constante de décomposition de la biomasse microbienne
    Dim Fbio As Double 'facteur de correction du C/N de la biomasse microbienne quand le N est limitant. Fixé à 1 pour l'instant
    Dim rstResidus As ADODB.Recordset
    ' prévu pour bilans nutriments journaliers, pas opérationnel
    Public Sub LisApportsOrga(TypeMorga As String)

    rstResidus = New ADODB.Recordset

    rstResidus.Open("ListResidus", DataBase_Cnn)
    rstResidus.MoveFirst
    Trouve = False

    While Not rstResidus.EOF And Not Trouve
        If rstResidus("TypeMorga") = TypeMorga Then
             Trouve = True
             Akres = rstResidus("Akres")
             Bkres = rstResidus("Bkres")
             AHres = rstResidus("AHres")
             BHres = rstResidus("BHres")
             AWB = rstResidus("AWB")
             BWB = rstResidus("BWB")
             Yres = rstResidus("Yres")
             Kbio = rstResidus("Kbio")
             Fbio = rstResidus("Fbio")
             Nrec = rstResidus("Nrec")
        End If
        rstResidus.MoveNext
    End While
    rstResidus.Close


    rstResidus = Nothing

    End Sub
    Public Function dNrec() As Double
    Return Nrec
    End Function
    Public Function dAkres() As Double
    Return Akres
    End Function
    Public Function dBkres() As Double
    Return Bkres
    End Function
    Public Function dAHres() As Double
    Return AHres
    End Function
    Public Function dBHres() As Double
    Return BHres
    End Function
    Public Function dAWB() As Double
    Return AWB
    End Function
    Public Function dBWB() As Double
    Return BWB
    End Function
    Public Function dYres() As Double
    Return Yres
    End Function
    Public Function dKbio() As Double
    Return Kbio
    End Function
    Public Function dFbio() As Double
    Return Fbio
    End Function
End Class
