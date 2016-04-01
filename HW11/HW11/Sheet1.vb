
Public Class Sheet1

    Private Sub Sheet1_Startup() Handles Me.Startup
        StockDataLst.AutoSetDataBoundColumnHeaders = True
        SetUpTheADOcomponents()
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
        LoadDistinctTickersInCBox()
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown
        DisconnectFromDB()
    End Sub

    Private Sub LoadDistinctTickersInCBox()
        Dim temp As String
        RunQueryAndSaveResultInDS("SELECT DISTINCT Ticker FROM StockMarket", "DistinctTickers")
        For Each row As DataRow In myDataSet.Tables("DistinctTickers").Rows
            temp = row("Ticker")
            TickerCBox.Items.Add(temp)
        Next
        MessageBox.Show("Ticker loaded.", "Information Window")
    End Sub

    Private Sub LatestClosingsBtn_Click(sender As Object, e As EventArgs) Handles LatestClosingsBtn.Click
        Dim maxDate As Date
        Dim myQuery As String

        maxDate = GetADate("SELECT max(Date) FROM StockMarket")
        Range("K3").Value = "Latest available closing date is " + maxDate.ToShortDateString()

        ClearTableInDS("ClosingsTbl")
        myQuery = String.Format("SELECT * FROM StockMarket WHERE Date = '{0}' ORDER BY Ticker",
                                maxDate.ToShortDateString)
        RunQueryAndSaveResultInDS(myQuery, "ClosingsTbl")
        StockDataLst.DataSource = myDataSet.Tables("ClosingsTbl")
        StockDataLst.TableStyle = "TableStyleDark8"
    End Sub

    Private Sub HistForTickerBtn_Click(sender As Object, e As EventArgs) Handles HistForTickerBtn.Click
        Dim myQuery As String
        ClearTableInDS("HistForTickerTbl")
        myQuery = String.Format("SELECT * FROM StockMarket WHERE Ticker = '{0}' ORDER BY Date DESC",
                                TickerCBox.SelectedItem)
        RunQueryAndSaveResultInDS(myQuery, "HistForTickerTbl")
        StockDataLst.DataSource = myDataSet.Tables("HistForTickerTbl")
        StockDataLst.TableStyle = "TableStyleMedium2"
    End Sub

    Private Sub TickerCBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TickerCBox.SelectedIndexChanged

    End Sub
End Class
