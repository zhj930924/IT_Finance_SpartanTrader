Module DBprocedures
    ' These are the ADO components
    Dim myConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection
    Dim myConnectionString As String = ""
    Dim myCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter
    Public myDataSet As DataSet = New DataSet

    Public Sub SetUpTheADOcomponents()
        Try
            'give to the command a connection
            myCommand.Connection = myConnection
            'give to the data adapter a command
            myDataAdapter.SelectCommand = myCommand
        Catch e As Exception
            MessageBox.Show("Ehi! SetUpTheADOcomponents failed: " + e.Message)
        End Try
    End Sub

    Public Sub ConnectToDB(connString As String)
        Try
            'set the connection string
            myConnection.ConnectionString = connString
            myConnection.Open()
        Catch e As Exception
            MessageBox.Show("Ehi! ConnectToDB failed: " + e.Message)
        End Try
    End Sub

    Public Sub DisconnectFromDB()
        Try
            myConnection.Close()
        Catch e As Exception
            MessageBox.Show("Ehi! DisconnectFromDB failed: " + e.Message)
        End Try
    End Sub

    Public Sub RunQueryAndSaveResultInDS(query As String, resultName As String)
        Try
            myCommand.CommandText = query
            myDataAdapter.Fill(myDataSet, resultName)
        Catch e As Exception
            MessageBox.Show("Ehi! RunQueryAndSaveResultInDS failed: " + e.Message)
        End Try
    End Sub

    Public Sub ClearTableInDS(tableName As String)
        Try
            If myDataSet.Tables.Contains(tableName) Then
                myDataSet.Tables(tableName).Clear()
            End If
        Catch e As Exception
            MessageBox.Show("Ehi! ClearTableInDS failed: " + e.Message)
        End Try
    End Sub

    Public Sub ExecuteNonQuery(query As String)
        Try
            myCommand.CommandText = query
            myCommand.ExecuteNonQuery()
        Catch e As Exception
            MessageBox.Show("Ehi! ExecuteNonQuery failed: " + e.Message)
        End Try
    End Sub

    Public Function GetADate(query As String) As Date
        Try
            myCommand.CommandText = query
            Return DateTime.Parse(myCommand.ExecuteScalar())
        Catch e As Exception
            MessageBox.Show("Opps! GetADate failed: " + e.Message + " 1/1/1 was returned.")
            Return DateTime.Parse("1/1/1")
        End Try
    End Function

End Module


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
