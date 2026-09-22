Option Compare Database
Public Sub GraphesJournaliers()
Dim MaBd As Database
Dim Db_Cnn As ADODB.Connection
Dim RequResuJ As New ADODB.Recordset
Dim RequA_Grapher As New ADODB.Recordset
Dim J As Integer
Dim Categ As String, CategPrecedent As String, ok As Boolean, TargetFile As String, NameSheet As String ',Idplot As String, CodeE As String

Set Db_Cnn = CurrentProject.Connection

Set MaBd = CurrentDb()

TargetFile = "D:\donneesFA\modelisation\Arise\CelsiusV31\test.xls"
'TableANumeroter.Open "Select * from IrrigParcellesElemBbey96_99PK order by idplotprov, CodeEau, CodeN, Bloc, dateIrrig", Db_Cnn, adOpenDynamic, adLockOptimistic
RequResuJ.Open "Select * from VisuResuJTacsy2 order by idPlotAn, YearS, Jul", Db_Cnn, adOpenDynamic, adLockOptimistic
RequResuJ.MoveFirst
J = 1

    While Not RequResuJ.EOF
        CategPrecedent = Categ
        Categ = RequResuJ!idPlotAn
        If Categ <> CategPrecedent Then
        RequA_Grapher.Open "Select * from VisuResuJTacsy2 where IdPlotAn='" & Categ & "'order by idPlotAn, YearS, Jul", Db_Cnn, adOpenDynamic, adLockOptimistic
        NameSheet = "pipo" & J
        
        DoCmd.Save acDefault, NameSheet
        
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, NameSheet, TargetFile, True
            
        J = J + 1
        End If
    RequResuJ.MoveNext
    
    Wend
    
End Sub