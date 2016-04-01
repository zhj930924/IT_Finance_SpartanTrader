Module ProcessAutomation

    Public Sub DownloadStaticData()

        DownloadInitialPositions()
        DownloadTransactionCosts()
        DownloadEnvironmentVariable()
        DownloadTickers()
        Globals.Dashboard.LoadCBoxes()
        ' we use 'get' to indicate that we are exracting data from the dataset, and 'download'
        ' to indicate that we are extracting data from the database
        initialCAccount = GetInitialCAccount()
        iRate = GetIRate()
        startDate = GetStartDate()
        endDate = GetEndDate()
        maxMargins = GetMaxMargins()
        TPVatStart = CalcTPVAtStart()

    End Sub

    ' -Homework 15----------------------------------------------
    Public Sub WarmStart()

        ClearOldDataInListObjects() 'new line
        CreateAndConnectTheADOObjects()
        If OpenDBConnection() = False Then
            Exit Sub ' because it could not connect
        End If
        DownloadStaticData()
        DownloadTeamData()
        currentDate = DownloadCurrentDate()
        DownloadPricesForOneDay(currentDate)
        SetupCharts()
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)

    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)

        'interestSLT = CalcInterestSLT(targetDate) <--- moved to Transaction.vb
        CAccount = CAccount + interestSLT
        margin = CalcMargin(targetDate)
        IP = CalcIPValue(targetDate)
        AP = CalcAPValue(targetDate)
        TPV = IP + AP + CAccount
        TaTPV = CalcTaTPV(targetDate)
        TE = TPV - TaTPV
        If TE > 0 Then TE = TE / 4 'If a gain then
        TEpercent = TE / TaTPV

    End Sub

    Public Sub DisplayFinancialMetrics(targetDate As Date)

        Try
            Globals.Dashboard.CAccountCell.Value = CAccount
            Globals.Dashboard.MarginCell.Value = margin
            Globals.Dashboard.MarginPercCell.Value = margin * 0.3
            Globals.Dashboard.maxMarginCell.Value = maxMargins

            Globals.Dashboard.InterestSLTCell.Value = interestSLT
            Globals.Dashboard.IPCell.Value = IP
            Globals.Dashboard.APCell.Value = AP

            Globals.Dashboard.TPVatStartCell.Value = TPVatStart
            Globals.Dashboard.TPVCell.Value = TPV
            Globals.Dashboard.TaTPVCell.Value = TaTPV
            Globals.Dashboard.TECell.Value = TE
            Globals.Dashboard.TEPercCell.Value = TEpercent

        Catch
            ' do nothing
        End Try

    End Sub

    Public Sub DownloadTeamData()

        CAccount = DownloadCAccount()
        DownloadAcquiredPositions()
        lastTransactionDate = DownloadLastTransactionDate(endDate)

    End Sub

    Public Sub SetupCharts()

        Globals.Dashboard.FillTPVTrackingTable()
        Globals.Dashboard.SetupTrackingChart()

    End Sub

    Public Sub ClearOldDataInListObjects()

        Globals.Markets.StockMarketLO.DataSource = Nothing
        Globals.Markets.OptionMarketLO.DataSource = Nothing
        Globals.Markets.SP500LO.DataSource = Nothing
        Globals.Portfolio.InitialPositionsLO.DataSource = Nothing
        Globals.Portfolio.AcquiredPositionsLO.DataSource = Nothing
        Globals.Environment.SettingsLO.DataSource = Nothing
        Globals.Environment.TransactionCostsLO.DataSource = Nothing

    End Sub

    Public Sub CalcAndDisplayFinancialMetrics(targetDate As Date)

        Globals.Dashboard.maxMarginCell.Value = maxMargins
        Globals.Dashboard.TPVatStartCell.Value = TPVatStart
        IP = CalcIPValue(targetDate)
        Globals.Dashboard.IPCell.Value = IP

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
