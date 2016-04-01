Imports Microsoft.Office.Tools.Ribbon

Public Class stRibbon

    Private Sub stRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ' this activates the custom ribbon at start
        Globals.Ribbons.stRibbon.RibbonUI.ActivateTabMso("TabAddIns")
    End Sub

    Private Sub AlphaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaBtn.Click
        AlphaBtn.Checked = True
        BetaBtn.Checked = False
        GammaBtn.Checked = False
        activeDB = "Alpha"
        WarmStart()
    End Sub

    Public Sub BetaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaBtn.Click
        AlphaBtn.Checked = False
        BetaBtn.Checked = True
        GammaBtn.Checked = False
        activeDB = "Beta"
        WarmStart()
    End Sub

    Public Sub GammaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles GammaBtn.Click
        AlphaBtn.Checked = False
        BetaBtn.Checked = False
        GammaBtn.Checked = True
        activeDB = "Gamma"
        WarmStart()
    End Sub

    Public Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click

    End Sub

    Public Sub InitialPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionsBtn.Click
        DownloadInitialPositions()
        ShowInitialPositions()
    End Sub

    Public Sub AcquiredPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AcquiredPositionsBtn.Click
        DownloadAcquiredPositions()
        ShowAcquiredPositions()
    End Sub

    Public Sub StockMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StockMktBtn.Click
        DownloadStockMarket()
        ShowStockMarket()
    End Sub

    Public Sub OptionMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionMktBtn.Click
        DownloadOptionMarket()
        ShowOptionMarket()
    End Sub

    Public Sub SP500Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles SP500Btn.Click
        DownloadStockIndex()
        ShowStockIndex()
    End Sub

    Public Sub SettingsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingsBtn.Click
        DownloadEnvironmentVariable()
        ShowEnvironmentVariable()
    End Sub

    Public Sub TCostsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TCostsBtn.Click
        DownloadTransactionCosts()
        ShowTransactionCosts()
    End Sub

    Public Sub QuitBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitBtn.Click
        stQuit()
    End Sub

End Class

Module ProcessAutomation

    Public Sub WarmStart()
        CreateAndConnectTheADOObjects()
        If OpenDBConnection() = False Then
            Exit Sub ' because it could not connect
        End If
    End Sub

    Public Sub stStart()
        ' Turn off formular bar to create more space
        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        ' Show the dashboard to us
        Globals.Dashboard.Activate()
        ' Click the beta button in the ribbon
        Globals.Ribbons.stRibbon.BetaBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub stQuit()
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.DisplayFormulaBar = True
        Globals.ThisWorkbook.Application.Quit()
    End Sub

End Module

Module DBUtilities

    'This Module contains procedures for managing DB connections and manipulating data
    '-- ADO-related objects (also global variables)

    Dim myConnection As SqlClient.SqlConnection
    Dim myCommand As SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet
    Dim myDataTable As DataTable
    Dim mySQLString As String

    Public Sub CreateAndConnectTheADOObjects()
        Try
            'Create the connection and set the connection string
            myConnection = New SqlClient.SqlConnection
            'Create the command and set the connection
            myCommand = New SqlClient.SqlCommand
            myCommand.Connection = myConnection
            'Create the data adapter and set the selectCommand
            myDataAdapter = New SqlClient.SqlDataAdapter
            myDataAdapter.SelectCommand = myCommand
            'Create the dataset
            myDataSet = New DataSet
        Catch e As Exception
            MessageBox.Show("Ehi! CreateAndConnectTheADOObjects failed: " + e.Message)
        End Try
    End Sub

    Public Function OpenDBConnection() As Boolean
        Select Case activeDB
            Case "Alpha"
                myConnection.ConnectionString = "Data Source = f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog = HedgeTournamentALPHA; Integrated Security = True"
            Case "Beta"
                myConnection.ConnectionString = "Data Source = f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog = HedgeTournamentBETA; Integrated Security = True"
            Case "Gamma"
                myConnection.ConnectionString = "Data Source = f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog = HedgeTournamentGAMMA; Integrated Security = True"
        End Select
        Try
            myConnection.Open()
            Return True ' true = success
        Catch myException As Exception
            MessageBox.Show("I am calling, but the DB is not responding, Dave. " + myException.Message,
                            "Connection problem", MessageBoxButtons.OK)
            Return False
        End Try

    End Function

    Public Sub DownloadInitialPositions()
        DownloadTableUsingSQL("Select * from InitialPosition order by symbol", "InitialPositionTable")
    End Sub

    Public Sub ShowInitialPositions()
        Globals.Portfolio.Activate()
        Globals.Portfolio.InitialPositionsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.InitialPositionsLO.DataSource = myDataSet.Tables("InitialPositionTable")
    End Sub

    Public Sub DownloadTableUsingSQL(mySQL As String, NameofTheResultTable As String)
        ClearDataSetTable(NameofTheResultTable)
        myCommand.CommandText = mySQL
        Try
            myDataAdapter.Fill(myDataSet, NameofTheResultTable)
        Catch myException As Exception
            MessageBox.Show("Dave, I could not download " + NameofTheResultTable + " using " + mySQL + ". " +
                            "No corrective action was taken. Maybe this will help: " + myException.Message,
                            "Likely SQL problem.", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub ClearDataSetTable(TableToClear As String)
        If myDataSet.Tables.Contains(TableToClear) Then
            myDataSet.Tables(TableToClear).Clear()
        End If
    End Sub

    Public Sub DownloadAcquiredPositions()
        DownloadTableUsingSQL("Select * from " + portfolioTableName + " order by symbol", portfolioTableName)
    End Sub

    Public Sub ShowAcquiredPositions()
        Globals.Portfolio.Activate()
        Globals.Portfolio.AcquiredPositionsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.AcquiredPositionsLO.DataSource = myDataSet.Tables(portfolioTableName)
    End Sub

    Public Sub DownloadStockMarket()
        DownloadTableUsingSQL("Select * from StockMarket", "StockMarketTable")
    End Sub

    Public Sub ShowStockMarket()
        Globals.Markets.Activate()
        Globals.Markets.StockMarketLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.StockMarketLO.DataSource = myDataSet.Tables("StockMarketTable")
    End Sub

    Public Sub DownloadOptionMarket()
        DownloadTableUsingSQL("Select * from OptionMarket", "OptionMarketTable")
    End Sub

    Public Sub ShowOptionMarket()
        Globals.Markets.Activate()
        Globals.Markets.OptionMarketLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.OptionMarketLO.DataSource = myDataSet.Tables("OptionMarketTable")
    End Sub

    Public Sub DownloadStockIndex()
        DownloadTableUsingSQL("Select * from StockIndex order by date desc", "StockIndexTable")
    End Sub

    Public Sub ShowStockIndex()
        Globals.Markets.Activate()
        Globals.Markets.SP500LO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.SP500LO.DataSource = myDataSet.Tables("StockIndexTable")
    End Sub

    Public Sub DownloadEnvironmentVariable()
        DownloadTableUsingSQL("Select * from EnvironmentVariable", "EnvironmentVariableTable")
    End Sub

    Public Sub ShowEnvironmentVariable()
        Globals.Environment.Activate()
        Globals.Environment.SettingsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Environment.SettingsLO.DataSource = myDataSet.Tables("EnvironmentVariableTable")
    End Sub

    Public Sub DownloadTransactionCosts()
        DownloadTableUsingSQL("Select * from TransactionCost", "TransactionCostTable")
    End Sub

    Public Sub ShowTransactionCosts()
        Globals.Environment.Activate()
        Globals.Environment.TransactionCostsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Environment.TransactionCostsLO.DataSource = myDataSet.Tables("TransactionCostTable")
    End Sub


End Module

Module stGlobals

    Public activeDB As String = ""
    Public teamID As String = "30"
    Public portfolioTableName As String = "PortfolioTeam" + teamID

End Module
