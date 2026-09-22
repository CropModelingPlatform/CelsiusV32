Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim ActiveStressH As Boolean 'vrai: stress hydrique actif sur variables  plante (après levée)
Dim ActivestressN As Boolean 'vrai: stress hydrique actif sur variables  plante (après levée)
Dim simlevee As Boolean 'vrai: simulation germination et levee ; faux levee forcée à date indiquée dans tecperCrop
Dim EcritDResus As Boolean 'vrai écriture des résultats journaliers dans "OutputD"
Dim Trouve As Boolean
Dim CorAlti As Boolean 'si vrai correction des températures en fonction de la différence d'altitue dentre station clim et unité simulée
Dim CyberST As Boolean 'si vrai activation de la gestion technique automatique (Cyber Système Technique...)
Dim TypeNPKstress As Integer 'choix de la méthode de calcul des stress nutritionnels 1: bilan saisonnier; 2: bilan journalier
Dim CodeDevelop As String 'option de calcul des stades phenos
Dim CCYNo As Boolean 'si vrai utilisation de scénarios de changement clim par méthode des deltas et lecture du scénario lors de lecture dataclim selon
                      ' code lu dans simUnitList (si faux le code n'est pas lu)

Dim rstOptionsModel As ADODB.Recordset
'module lisant les paramètres de choix des options de simulation
'VERSION 3 du 8-10 nov 2017'
' modif FA du 30/08/18 pour scenarios CC: lecture nouveau champ CCYNo
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026



Sub LisOptionsModel(DataBase_Cnn As ADODB.Connection, idOptionsModel As Integer)
Set rstOptionsModel = New ADODB.Recordset
Trouve = False

rstOptionsModel.Open "OptionsModel", DataBase_Cnn
rstOptionsModel.MoveFirst
While Not rstOptionsModel.EOF And Not Trouve
    If rstOptionsModel!idCodModel = idOptionsModel Then
        
        Trouve = True
        ActiveStressH = rstOptionsModel!ActiveWstress
        ActivestressN = rstOptionsModel!ActiveNstress
        simlevee = rstOptionsModel!simlevee
        EcritDResus = rstOptionsModel!EcritDResus
        CorAlti = rstOptionsModel!CorrigAlti
        CyberST = rstOptionsModel!CyberST
        TypeNPKstress = rstOptionsModel!TypeNPKstress
        CodeDevelop = rstOptionsModel!CodeDevelop
        CCYNo = rstOptionsModel!CCYNo
        'introduire ici option mauvaises herbes
    End If
    rstOptionsModel.MoveNext
Wend
'fermeture table et libération mémoire de l'objet
If CyberST And simlevee = False Then
    MsgBox ("CyberST vrai et Simlevee faux ! Simlevee sera reglé sur True")
    simlevee = True
End If

rstOptionsModel.Close
Set rstOptionsModel = Nothing
End Sub
Property Get bActiveStressH() As Boolean
bActiveStressH = ActiveStressH
End Property
Property Get bActiveStressN() As Boolean
bActiveStressN = ActivestressN
End Property
Property Get bSimLevee() As Boolean
bSimLevee = simlevee
End Property
Property Get bEcritDresus() As Boolean
bEcritDresus = EcritDResus
End Property
Property Get bCorAlti() As Boolean
bCorAlti = CorAlti
End Property
Property Get bCyberST() As Boolean
bCyberST = CyberST
End Property
Property Get nTypeNPKstress() As Integer
nTypeNPKstress = TypeNPKstress
End Property
Property Get sCodeDevelop() As String
sCodeDevelop = CodeDevelop
End Property
Property Get bCCYNo() As Boolean
bCCYNo = CCYNo
End Property