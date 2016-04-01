Module DBUtilities

    'This Module contains procedures for managing DB connections and manipulating data
    '-- ADO-related objects (also global variables)

    Public Function DownloadDividend(ticker As String, targetDate As Date)

        Dim temp As String = "0"
        Dim mySql As String = ""

        ' last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(ticker) Then
            mySql = "Select dividend from StockMarket where ticker = '" + ticker + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy Batplane! I could not retrieve the dividend for " + ticker + ". This is the query you created " +
                            mySql + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Double.Parse(temp)

    End Function

    Public Sub ExecuteNonQuery(SQLString As String)

        Try
            myCommand.CommandText = SQLString
            myCommand.ExecuteNonQuery()
        Catch myException As Exception
            MessageBox.Show("You query failed, Dave." +
                            "Maybe this will help: " + myException.Message,
                            "Likely SQL problem", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub DownloadTransactionQueue(team As String)
        DownloadTableUsingSQL("Select * from TransactionQueue where teamID = '" + team + "' order by rowID desc",
                              "TransactionQueueTable")
    End Sub

    Public Sub ShowTransactionQueue()

        Globals.TransactionQueue.Activate()
        Globals.TransactionQueue.TransactionQueueLO.AutoSetDataBoundColumnHeaders = True
        Globals.TransactionQueue.TransactionQueueLO.TableStyle = "TableStyleDark8"
        Globals.TransactionQueue.TransactionQueueLO.DataSource = myDataSet.Tables("TransactionQueueTable")

    End Sub

    '-- Homework 16-------------------------------------------------------------------------------------

    Public Sub DownloadPricesForOneDay(targetDate As Date)

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If targetDate.Date <> lastPriceDownloadDate.Date Then

            Dim mySQL As String

            mySQL = "Select * from StockMarket where Date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "StockMarketOneDayTable")

            mySQL = "Select * from OptionMarket where Date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "OptionMarketOneDayTable")
            lastPriceDownloadDate = targetDate

        End If

    End Sub

    Public Function DownloadLastTransactionDate(targetDate As Date) As Date

        Dim temp As String = ""
        myCommand.CommandText = String.Format("Select max(date) from TransactionQueue where teamid = {0} and date <= '{1}'",
                                              teamID, targetDate.ToShortDateString())
        Try
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch myException As Exception
            MessageBox.Show("Last transaction not found. Set LastTransactionDate to StartDate ",
                            "Transaction Queue", MessageBoxButtons.OK)
            Return startDate
        End Try

    End Function

    Public Function DownloadCAccount() As Double

        Dim temp As String = ""
        myCommand.CommandText = "Select Units from " + portfolioTableName + " where Symbol = 'CAccount'"
        Try
            temp = myCommand.ExecuteScalar()
            Return Double.Parse(temp)
        Catch myException As Exception
            MessageBox.Show("Holy batmobile! I could not retrieve the CAccount. I reported $0. " +
                            "Maybe this will help: " + myException.Message,
                            "Likely SQL problem", MessageBoxButtons.OK)
            Return 0
        End Try

    End Function

    Public Function DownloadCurrentDate() As Date

        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where Name = 'CurrentDate'"
        Try
            temp = myCommand.ExecuteScalar()
            Globals.Dashboard.CurrentDateCell.Value = Date.Parse(temp).ToLongDateString()
            Return Date.Parse(temp)
        Catch myException As Exception
            Return currentDate
        End Try

    End Function

    '-------------------------------------------Homework 14---------------------------------------------

    Public Sub DownloadTickers()
        DownloadTableUsingSQL("Select distinct ticker from StockMarket order by ticker", "TickerTable")
    End Sub

    Public Function DownloadAsk(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        ' last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(symbol) Then
            mySql = "Select Ask from StockMarket where ticker = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Ask from OptionMarket where symbol = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy Batmobile! I could not retrieve the ask for " + symbol + ". This is the query you created " +
                            mySql + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Double.Parse(temp)

    End Function

    Public Function DownloadBid(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        ' last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(symbol) Then
            mySql = "Select Bid from StockMarket where ticker = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Bid from OptionMarket where symbol = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy Batgirl! I could not retrieve the bid for " + symbol + ". This is the query you created " +
                            mySql + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Double.Parse(temp)

    End Function

    '----------------- Homework 13 ----------------------------------------------------
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
        DownloadTableUsingSQL("Select * from StockMarket order by Date desc", "StockMarketTable")
    End Sub

    Public Sub ShowStockMarket()
        Globals.Markets.Activate()
        Globals.Markets.StockMarketLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.StockMarketLO.DataSource = myDataSet.Tables("StockMarketTable")
    End Sub

    Public Sub DownloadOptionMarket()
        DownloadTableUsingSQL("Select * from OptionMarket order by Date desc", "OptionMarketTable")
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
