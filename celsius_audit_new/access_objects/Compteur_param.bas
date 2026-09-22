Option Compare Database
'set of functions usefull for building cumulative  distribution curves after virtual experiments across series of climate data
Function Lance_param(nanFreq As Integer)
Dim strNomTable As String
strNomTable = "resuExpeVirtu"
Call AjoutCompteur(strNomTable, nanFreq)
'Call AjoutCompteur(strNomTable:="ResuExpeVirtu", NanAfreq:=21)
End Function
' ajout de compteurs pour classer des tables en fonction de plusieurs champs avant construction de graphiques de distribution cumulées

Function AjoutCompteur(strNomTable As String, NanAfreq As Integer) As Boolean

On Error GoTo TrappeErreur

Dim Db_Cnn As ADODB.Connection

Dim TabAnaFreq As New ADODB.Recordset
Dim J As Integer

Set Db_Cnn = CurrentProject.Connection


    ' Vérifie si une table du même nom existe déjà
If VerifierExistenceTable(strNomTable:="ResuExpeVirtu") = False Then
    MsgBox ("faut créer la table  " & vbCrLf & strNomTable & vbCrLf & "avant d'y mettre des données, banane flambée ! Fin !")
 GoTo Sortie
End If
Set TabAnaFreq = New ADODB.Recordset

TabAnaFreq.Open "SELECT * FROM ResuExpeVirtu Order by idweather,NomSC,idsoil,CodParamMulch,isem,rdtgrain", Db_Cnn, adOpenDynamic, adLockOptimistic

While Not TabAnaFreq.EOF
    For J = 1 To NanAfreq
    TabAnaFreq.Fields(10).Value = J
    TabAnaFreq.Update
    TabAnaFreq.MoveNext
    Next J
Wend
TabAnaFreq.Close
TabAnaFreq.Open "SELECT * FROM ResuExpeVirtu Order by idweather,NomSC,idsoil,CodParamMulch,isem,Biotot", Db_Cnn, adOpenDynamic, adLockOptimistic
While Not TabAnaFreq.EOF
    For J = 1 To NanAfreq
    TabAnaFreq!CompteurBiomasse.Value = J
    TabAnaFreq.Update
    TabAnaFreq.MoveNext
    Next J
Wend
TabAnaFreq.Close
        
Set TabAnaFreq = Nothing
    

Sortie:
 Set TabAnaFreq = Nothing
 Set Db_Cnn = Nothing
 Exit Function

TrappeErreur:
 MsgBox Err.Description
 Resume Sortie

End Function


Function VerifierExistenceTable(strNomTable As String) As Boolean ' Verifie si une table existe dans la base de données courante
    Dim tblTable As DAO.TableDef

    VerifierExistenceTable = False
    For Each Tb In CurrentDb.TableDefs
        If Tb.Name = strNomTable Then
        VerifierExistenceTable = True
    Exit For
        End If
    Next
End Function