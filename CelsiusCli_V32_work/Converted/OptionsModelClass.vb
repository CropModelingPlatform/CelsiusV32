Option Strict Off
Option Explicit Off
Imports System
Public Class OptionsModelClass
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
    Public Sub LisOptionsModel(DataBase_Cnn As ADODB.Connection, idOptionsModel As Integer)
    rstOptionsModel = New ADODB.Recordset
    Trouve = False

    rstOptionsModel.Open("OptionsModel", DataBase_Cnn)
    rstOptionsModel.MoveFirst
    While Not rstOptionsModel.EOF And Not Trouve
        If rstOptionsModel("idCodModel") = idOptionsModel Then

            Trouve = True
            ActiveStressH = rstOptionsModel("ActiveWstress")
            ActivestressN = rstOptionsModel("ActiveNstress")
            simlevee = rstOptionsModel("simlevee")
            EcritDResus = rstOptionsModel("EcritDResus")
            CorAlti = rstOptionsModel("CorrigAlti")
            CyberST = rstOptionsModel("CyberST")
            TypeNPKstress = rstOptionsModel("TypeNPKstress")
            CodeDevelop = rstOptionsModel("CodeDevelop")
            CCYNo = rstOptionsModel("CCYNo")
            'introduire ici option mauvaises herbes
        End If
        rstOptionsModel.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet
    If CyberST And simlevee = False Then
    Console.Error.WriteLine(("CyberST vrai et Simlevee faux ! Simlevee sera reglé sur True"))
        simlevee = True
    End If

    rstOptionsModel.Close
    rstOptionsModel = Nothing
    End Sub
    Public Function bActiveStressH() As Boolean
    Return ActiveStressH
    End Function
    Public Function bActiveStressN() As Boolean
    Return ActivestressN
    End Function
    Public Function bSimLevee() As Boolean
    Return simlevee
    End Function
    Public Function bEcritDresus() As Boolean
    Return EcritDResus
    End Function
    Public Function bCorAlti() As Boolean
    Return CorAlti
    End Function
    Public Function bCyberST() As Boolean
    Return CyberST
    End Function
    Public Function nTypeNPKstress() As Integer
    Return TypeNPKstress
    End Function
    Public Function sCodeDevelop() As String
    Return CodeDevelop
    End Function
    Public Function bCCYNo() As Boolean
    Return CCYNo
    End Function
End Class
