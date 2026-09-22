Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
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

Sub LisApportsOrga(TypeMorga As String)

Set rstResidus = New ADODB.Recordset

rstResidus.Open "ListResidus", DataBase_Cnn
rstResidus.MoveFirst
Trouve = False

While Not rstResidus.EOF And Not Trouve
    If rstResidus!TypeMorga = TypeMorga Then
         Trouve = True
         Akres = rstResidus!Akres
         Bkres = rstResidus!Bkres
         AHres = rstResidus!AHres
         BHres = rstResidus!BHres
         AWB = rstResidus!AWB
         BWB = rstResidus!BWB
         Yres = rstResidus!Yres
         Kbio = rstResidus!Kbio
         Fbio = rstResidus!Fbio
         Nrec = rstResidus!Nrec
    End If
    rstResidus.MoveNext
Wend
rstResidus.Close


Set rstResidus = Nothing

End Sub
Property Get dNrec() As Double
dNrec = Nrec
End Property
Property Get dAkres() As Double
dAkres = Akres
End Property
Property Get dBkres() As Double
dBkres = Bkres
End Property
Property Get dAHres() As Double
dAHres = AHres
End Property
Property Get dBHres() As Double
dBHres = BHres
End Property
Property Get dAWB() As Double
dAWB = AWB
End Property
Property Get dBWB() As Double
dBWB = BWB
End Property
Property Get dYres() As Double
dYres = Yres
End Property
Property Get dKbio() As Double
dKbio = Kbio
End Property
Property Get dFbio() As Double
dFbio = Fbio
End Property