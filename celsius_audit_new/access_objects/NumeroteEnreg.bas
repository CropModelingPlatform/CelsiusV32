Option Compare Database
Public Sub NumeroteEnreg()
Dim MaBd As Database
Dim Db_Cnn As ADODB.Connection
Dim TableANumeroter As New ADODB.Recordset

Dim J As Integer
Dim Idplot As String, CodeE As String, CodeN As Variant, Bl As Variant, datI As Variant, Categ As String

Set Db_Cnn = CurrentProject.Connection

Set MaBd = CurrentDb()


'TableANumeroter.Open "Select * from IrrigParcellesElemBbey96_99PK order by idplotprov, CodeEau, CodeN, Bloc, dateIrrig", Db_Cnn, adOpenDynamic, adLockOptimistic
TableANumeroter.Open "Select * from CondInit order by idexpe, icbl", Db_Cnn, adOpenDynamic, adLockOptimistic
TableANumeroter.MoveFirst
J = 1

    While Not TableANumeroter.EOF
        CategPrecedent = Categ
        'Idplot = TableANumeroter!idplotprov
        Idplot = TableANumeroter!idexpe
        'CodeE = TableANumeroter!CodeEau
        'CodeN = TableANumeroter!CodeN
        'Bl = TableANumeroter!Bloc
        'categ = Idplot & CodeE & CodeN & Bl
        Categ = Idplot
        If Categ = CategPrecedent Or J = 1 Then
            'TableANumeroter!NumIrrig.Value = J
            TableANumeroter!NumInit.Value = J
            TableANumeroter.Update
            TableANumeroter.MoveNext
            J = J + 1
        Else
        J = 1
        End If
               
    Wend
    
 
    

TableANumeroter.Close

        
Set TableANumeroter = Nothing
End Sub