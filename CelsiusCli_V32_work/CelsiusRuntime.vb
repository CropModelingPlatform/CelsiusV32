Imports System
Imports System.Data
Imports System.Data.SQLite

Module CelsiusRuntime
    Public Sub Run(sqlitePath As String)
        Using sqlite = New SQLiteConnection($"Data Source={sqlitePath};Version=3;")
            sqlite.Open()
            EnsureOutputTables(sqlite)

            Dim outputDb = New ADODB.Connection()
            outputDb.Open($"Data Source={sqlitePath};Version=3;")

            Try
                PrincipalRunner.Run(sqlite, outputDb)
            Finally
                outputDb.Close()
            End Try
        End Using
    End Sub

    Private Sub EnsureOutputTables(connection As SQLiteConnection)
        Using command = connection.CreateCommand()
            command.CommandText = "
CREATE TABLE IF NOT EXISTS [OutputSynt] (
    [Idsim] TEXT NOT NULL,
    [JulPheno1_1] INTEGER,
    [JulPheno1_2] INTEGER,
    [JulPheno1_3] INTEGER,
    [JulPheno1_4] INTEGER,
    [JulPheno1_5] INTEGER,
    [JulPheno1_6] INTEGER,
    [death_day] INTEGER,
    [Biom(nrec)] REAL,
    [Grain(nrec)] REAL,
    [LAI] REAL,
    [SigmaSimEsol] REAL,
    [SigmaSimDr] REAL,
    [SigmaSimDrprofmax] REAL,
    [StockSol_init] REAL,
    [StockSol_final] REAL,
    [SigmaSimEmulch] REAL,
    [SigmaSimRuis] REAL,
    [SigmaTranspiMC] REAL,
    [SigmaSimPluM] REAL,
    [Ngrain] INTEGER,
    [P1grain] REAL,
    [Vitmoy] REAL,
    [iplt] INTEGER,
    [Nbsemis] INTEGER,
    [JulPheno2_1] INTEGER,
    [JulPheno2_2] INTEGER,
    [JulPheno2_3] INTEGER,
    [JulPheno2_4] INTEGER,
    [JulPheno2_5] INTEGER,
    [JulPheno2_6] INTEGER,
    [death_day_2] INTEGER,
    [Biom_nrec_2] REAL,
    [Grain_nrec_2] REAL,
    [LAI_2] REAL,
    [Ngrain_2] INTEGER,
    [P1grain_2] REAL,
    [Vitmoy_2] REAL,
    [iplt_2] INTEGER,
    [Nbsemis_2] INTEGER,
    [nbjContHlev] INTEGER,
    [stockNsol] REAL,
    [RuisEtrange] INTEGER,
    [SigmaCultEsol] REAL
);

CREATE TABLE IF NOT EXISTS [OutputD_1] (
    [IdSimJ] TEXT,
    [idSim] TEXT,
    [idweather] TEXT,
    [YearS] INTEGER,
    [idTech_Com] TEXT,
    [IdCultivar] TEXT,
    [DAP] INTEGER,
    [Jul] INTEGER,
    [Tmax] REAL,
    [Tmin] REAL,
    [Dvst] REAL,
    [Currestge] INTEGER,
    [Cropsta] INTEGER,
    [SomT] REAL,
    [LAI] REAL,
    [Biom] REAL,
    [grain] REAL,
    [Zrac] REAL,
    [Stsurf] REAL,
    [Esol] REAL,
    [Stnonrac] REAL,
    [Strac] REAL,
    [Transpi] REAL,
    [Drprofmax] REAL,
    [StockProf] REAL,
    [StockTot] REAL,
    [Plu] REAL,
    [Stockmes] REAL,
    [Emulch] REAL,
    [Ruis] REAL,
    [Qpaillis] REAL,
    [P1grain] REAL,
    [Etp] REAL,
    [TpotMC] REAL,
    [DrainageON] INTEGER,
    [ConcNsol] REAL,
    [Dr] REAL,
    [perteNDrain] REAL,
    [Nuptake] REAL,
    [SigmaNuptake] REAL,
    [perteGaz] REAL,
    [AppMinNj] REAL,
    [AppOrgNj] REAL,
    [Navail] REAL,
    [NUPTtarget] REAL,
    [NRF] REAL,
    [WSfactH] REAL,
    [WSfact] REAL,
    [TurfacH] REAL,
    [Turfac] REAL,
    [Stger] REAL,
    [ContrainteHlevee] INTEGER,
    [Competition] INTEGER,
    [CompFac] REAL,
    [raint] REAL,
    [SignalFletrissement] REAL,
    [Pfactor] REAL,
    [PfactorMC] REAL
);

CREATE TABLE IF NOT EXISTS [OutputD_2] AS SELECT * FROM [OutputD_1] WHERE 1 = 0;
"
            command.ExecuteNonQuery()
        End Using
    End Sub
End Module

Module PrincipalRunner
    Public Sub Run(inputDb As SQLiteConnection, outputDb As ADODB.Connection)
        Dim tabSynt As New ADODB.Recordset()
        Dim simUnits As New ADODB.Recordset()
        Dim compteSim As Long
        Dim simCtrl As New SimulationControlClass()
        Dim totalSimulations As Integer
        Dim lastPercent As Integer = -1

        Using countCommand As New SQLiteCommand("SELECT COUNT(*) FROM SimUnitList", inputDb)
            totalSimulations = Convert.ToInt32(countCommand.ExecuteScalar())
        End Using

        simUnits.Open("SELECT * FROM SimUnitList ORDER BY ChampTri", outputDb)
        tabSynt.Open("OutputSynt", outputDb, , ADODB.adLockOptimistic)
        ClearRecordset(tabSynt)

        Console.WriteLine("Progression VB.NET : 0%")
        lastPercent = 0
        While Not simUnits.EOF
            Codesuite = DbInt(simUnits("Codesuite"))
            simCtrl.ReadParameters(simUnits, outputDb, compteSim)
            simCtrl.Simulation()
            simCtrl.SortieSynthesis(tabSynt)
            simCtrl.EcritDresu(outputDb, compteSim)
            simCtrl.MemoEtatFinal()

            simUnits.MoveNext()
            compteSim += 1
            Dim percent = If(totalSimulations = 0, 100, CInt(Math.Floor(compteSim * 100.0 / totalSimulations)))
            If percent > lastPercent Then
                Console.WriteLine($"Progression VB.NET : {percent}% ({compteSim}/{totalSimulations})")
                lastPercent = percent
            End If
        End While

        simUnits.Close()
        tabSynt.Close()
    End Sub

    Private Sub ClearRecordset(recordset As ADODB.Recordset)
        recordset.MoveFirst()
        While Not recordset.EOF
            recordset.Delete()
            recordset.MoveNext()
        End While
    End Sub
End Module

Module CelsiusMath
    Public Function Inverse(x As Double) As Double
        If x = 0 Then
            Throw New InvalidOperationException("constantes thermiques nulles !")
        End If

        Return 1 / x
    End Function

    Public Function Max(x As Double, y As Double) As Double
        Return Math.Max(x, y)
    End Function

    Public Function Min(x As Double, y As Double) As Double
        Return Math.Min(x, y)
    End Function
End Module
