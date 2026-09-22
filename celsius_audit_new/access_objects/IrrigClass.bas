Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Dim Irr(0 To 731) As Double 'VX irrigations journalières en mm chaque jour depuis le 1er jour de simulation (indice 1)
'Dim dateIrr(0 To 731)  'VX date des irrigations en jour depuis le 1er jour de simul
Dim NbreIrrig As Integer
Dim rstDataIrr As ADODB.Recordset
' reste des choses à vérifier, cf commentaires
 'VERSION 3 du 8-10 nov 2017'
 ' verifiée et nettoyée FA le 07/05/2026

Sub LisIrrD(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass, GestionTechnique As GestionTechniqueClass)
' pourquoi passer Gestiontechnique ? Inutile à priori...
Dim AnIrrig As Integer
Dim jourIrrig As Integer
Dim irrig As Double
Dim Joursim As Integer
Dim ErrFa As Boolean
Dim Ndyear1 As Integer
Dim msg As String
On Error GoTo Err_LisIrrD

If SimUnit.bSY_Bissextile Then Ndyear1 = 366 Else Ndyear1 = 365
Set rstDataIrr = New ADODB.Recordset
'Lecture table "Irrigation" à modifier si nécessaire

rstDataIrr.Open "SELECT * FROM Irrigation_List where IdTech_Com='" & SimUnit.sIdTec & "' Order by DateIrrig", DataBase_Cnn, adOpenDynamic

rstDataIrr.MoveFirst


While Not rstDataIrr.EOF
    

    AnIrrig = rstDataIrr!YearIrrig
    jourIrrig = rstDataIrr!JourYrIrrig
    irrig = rstDataIrr!Irrigation
    
    If AnIrrig = SimUnit.nStartYear Then
        Joursim = jourIrrig - SimUnit.nStartDay + 1
    Else
        Joursim = jourIrrig + Ndyear1 - SimUnit.nStartDay + 1
    End If
    If Joursim < 0 Then
        msg = "Attention, calendrier des irrigations incohérent avec calendrier de simulation"
        GoTo ErrFA_LisIrrD
    Else
        Irr(Joursim) = irrig
        NbreIrrig = NbreIrrig + 1
    End If
    rstDataIrr.MoveNext
Wend



If NbreIrrig = 0 Then
    msg = "error in irrigation table, no irrigation found during simulation period"
    GoTo ErrFA_LisIrrD
End If

rstDataIrr.Close
Set rstDataIrr = Nothing



Exit_LisIrrD:
    Exit Sub
ErrFA_LisIrrD:
    ErrFa = True
    MsgBox ("ERREUR ! " & msg)
    Resume Exit_LisIrrD
Err_LisIrrD:
    MsgBox Err.Description
    Resume Exit_LisIrrD
End Sub
Property Get dIrr(J As Integer) As Double
dIrr = Irr(J)
End Property