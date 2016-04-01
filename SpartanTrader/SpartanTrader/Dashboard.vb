
Public Class Dashboard

    Private Sub Sheet1_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub SellShortStockBtn_Click(sender As Object, e As EventArgs) Handles SellShortStockBtn.Click

        myTransaction.Clear()
        myTransaction.action = "SellShort"
        myTransaction.trType = "SellShort"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeStockTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub
    Private Sub SellStockBtn_Click(sender As Object, e As EventArgs) Handles SellStockBtn.Click

        myTransaction.Clear()
        myTransaction.action = "Sell"
        myTransaction.trType = "Sell"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeStockTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub
    Private Sub CashDivBtn_Click(sender As Object, e As EventArgs) Handles CashDivBtn.Click

        myTransaction.Clear()
        myTransaction.action = "CashDiv"
        myTransaction.typeOfPrice = "Div"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeStockTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub

    Private Sub ExecuteStockTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteStockTransactionBtn.Click

        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeStockTransactionProperties()
            myTransaction.ExecuteTransaction()
            HighlightTransaction()
            CalcFinancialMetrics(currentDate)
            myTransaction.DisplayTransactionData()
            DisplayFinancialMetrics(currentDate)
        Else
            MessageBox.Show("I cannot do this for you, Dave. Stock input not valid.")
        End If

    End Sub

    Public Sub HighlightTransaction()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(0, 255, 0)
    End Sub

    Public Sub ClearTransactionHighlight()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(255, 255, 255)
    End Sub

    '-- Homework 16----------------------------------------------------------------------------
    Private Sub BuyStockBtn_Click(sender As Object, e As EventArgs) Handles BuyStockBtn.Click

        myTransaction.Clear()
        myTransaction.action = "Buy"
        myTransaction.trType = "Buy"
        myTransaction.typeOfPrice = "Ask"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeStockTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub

    Public Sub LoadCBoxes()

        TickerCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickerTable").Rows
            TickerCBox.Items.Add(myRow("Ticker").ToString().Trim())
        Next
        TickerCBox.Text = "Select Ticker"

    End Sub

    Public Sub SetupTrackingChart()

        ' format the chart
        TrackingChart.ChartType = Excel.XlChartType.xlLine
        TrackingChart.ChartStyle = 8
        TrackingChart.ApplyLayout(3)
        TrackingChart.HasTitle = False
        TrackingChart.HasLegend = True

        ' format the y axis as $
        Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"
        y.MinimumScaleIsAuto = False
        y.MaximumScaleIsAuto = True

        ' format the x axis as dates
        Dim x As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        TrackingChart.SetSourceData(TrackingChartLO.Range)
        Dim s As Excel.SeriesCollection = TrackingChart.SeriesCollection
        s(0).Format.Line.Weight = 2
        s(0).Format.Line.ForeColor.RGB = System.Drawing.Color.DarkOrange
        s(1).Format.Line.Weight = 2
        s(1).Format.Line.ForeColor.RGB = System.Drawing.Color.Gray
        s(2).Format.Line.Weight = 2
        s(2).Format.Line.ForeColor.RGB = System.Drawing.Color.DarkBlue

    End Sub

    Public Sub FillTPVTrackingTable()

        If myDataSet.Tables.Contains("TPVTrackingTable") Then
            myDataSet.Tables("TPVTrackingTable").Clear()
        Else
            ' create the table
            myDataSet.Tables.Add("TPVTrackingTable")
            myDataSet.Tables("TPVTrackingTable").Columns.Add("Date", GetType(Date))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TaTPV", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("NoHedge", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TPV", GetType(Double))
        End If

        ' fill it
        Dim tempTaTPV, tempTPV, tempNoHedge As Double
        Dim targetdate As Date

        For i As Integer = 14 To 0 Step -1
            targetdate = currentDate.AddDays(-i)
            If targetdate >= startDate Then
                DownloadPricesForOneDay(targetdate)
                tempTPV = CalcTPV(targetdate)
                tempTaTPV = CalcTaTPV(targetdate)
                tempNoHedge = CalcTPVNoHedge(targetdate)
                UpdateTPVTrackingTable(targetdate, tempTPV, tempTaTPV, tempNoHedge)
            End If
        Next

        TrackingChartLO.DataSource = myDataSet.Tables("TPVTrackingTable")

    End Sub

    Public Sub UpdateTPVTrackingTable(targetdate As Date, tpvInput As Double, tatpvInput As Double, noHedgeInput As Double)

        For Each myRow As DataRow In myDataSet.Tables("TPVTrackingTable").Rows
            If myRow("Date") = targetdate.ToShortDateString Then
                myRow("TPV") = tpvInput
                myRow("TaTPV") = tatpvInput
                myRow("NoHedge") = noHedgeInput
                Return
            End If
        Next
        myDataSet.Tables("TPVTrackingTable").Rows.Add(targetdate, tatpvInput, noHedgeInput, tpvInput)

        Try
            ' this line sts the scale of the chart for better viewing
            Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
            y.MinimumScale = Math.Truncate((FindMinInTPVTrackingTable() / 10000000)) * 10000000
        Catch
            ' skip screen refresh errors
        End Try

    End Sub

    Public Function FindMinInTPVTrackingTable() As Integer

        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables("TPVTrackingTable").Rows
            tempMin = Math.Min(myRow("TPV"), tempMin)
            tempMin = Math.Min(myRow("TaTPV"), tempMin)
            tempMin = Math.Min(myRow("NoHedge"), tempMin)
        Next
        Return tempMin
    End Function

    Private Sub TrackingChartLO_Change(targetRange As Excel.Range, changedRanges As ListRanges) Handles TrackingChartLO.Change

    End Sub

    Private Sub BuyOptionBtn_Click(sender As Object, e As EventArgs) Handles BuyOptionBtn.Click

    End Sub

    Private Sub ExecuteOptionTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteOptionTransactionBtn.Click

    End Sub

    Private Sub CAccountCell_Change(Target As Excel.Range) Handles CAccountCell.Change

    End Sub
End Class
