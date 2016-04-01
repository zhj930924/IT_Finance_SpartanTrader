﻿Module DataSetUtilities

    Public Function GetDividend(ticker As String, targetDate As Date) As Double

        If IsAStock(ticker) Then

            If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
                targetDate = targetDate.AddDays(-1)
            End If

            If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
                targetDate = targetDate.AddDays(-2)
            End If

            If targetDate.Date <> lastPriceDownloadDate.Date Then
                Return DownloadDividend(ticker, targetDate)
            End If

            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
                If myRow("Ticker").Trim() = ticker And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Dividend"))
                End If
            Next
        End If
        MessageBox.Show("Holy Batbelt! Could not find the dividend for " + ticker + ". Returned 0.")
        Return 0

    End Function


    Public Function GetTCostCoefficient(symbol As String, trType As String) As Double

        Dim tempTypeOfSecurity As String

        If IsAStock(symbol) Then
            tempTypeOfSecurity = "Stock"
        Else
            tempTypeOfSecurity = "Option"
        End If

        For Each myRow As DataRow In myDataSet.Tables("TransactionCostTable").Rows
            If myRow("SecurityType").Trim() = tempTypeOfSecurity And myRow("TransactionType").Trim() = trType Then
                Return Double.Parse(myRow("CostCoeff"))
            End If
        Next
        MessageBox.Show("Holy Batwig! Could not find the tr. cost for " + symbol + ". Returned 0.")
        Return 0

    End Function


    '---- Homework 15----------------------------------------------------------------------
    Public Function GetInitialCAccount() As Double

        Dim tempName, tempValue As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTable").Rows
            tempName = myRow("Name").ToString().Trim()
            If tempName = "CAccount" Then
                tempValue = myRow("Value").ToString().Trim()
                Return Double.Parse(tempValue)
            End If
        Next
        MessageBox.Show("Holy Batarang! Could not find 'CAccount'. Returned 0")
        Return 0

    End Function

    Public Function GetStartDate() As Date

        Dim tempName, tempValue As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTable").Rows
            tempName = myRow("Name").ToString().Trim()
            If tempName = "StartDate" Then
                tempValue = myRow("Value").ToString().Trim()
                Return Date.Parse(tempValue)
            End If
        Next
        MessageBox.Show("Holy Batarang! Could not find 'StartDate'. Returned nothing.")
        Return Nothing

    End Function

    Public Function GetEndDate() As Date

        Dim tempName, tempValue As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTable").Rows
            tempName = myRow("Name").ToString().Trim()
            If tempName = "EndDate" Then
                tempValue = myRow("Value").ToString().Trim()
                Return Date.Parse(tempValue)
            End If
        Next
        MessageBox.Show("Holy Batarang! Could not find 'EndDate'. Returned nothing.")
        Return Nothing

    End Function

    Public Function GetIRate() As Double

        Dim tempName, tempValue As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTable").Rows
            tempName = myRow("Name").ToString().Trim()
            If tempName = "RiskFreeRate" Then
                tempValue = myRow("Value").ToString().Trim()
                Return Double.Parse(tempValue)
            End If
        Next
        MessageBox.Show("Holy Batarang! Could not find 'RiskFreeRate'. Returned 0.")
        Return 0

    End Function

    Public Function GetMaxMargins() As Double

        Dim tempName, tempValue As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTable").Rows
            tempName = myRow("Name").ToString().Trim()
            If tempName = "MaxMargins" Then
                tempValue = myRow("Value").ToString().Trim()
                Return Double.Parse(tempValue)
            End If
        Next
        MessageBox.Show("Holy Batarang! Could not find 'MaxMargins'. Returned 0.")
        Return 0

    End Function

    Public Function GetAsk(symbol As String, targetDate As Date) As Double

        If targetDate.Date <> currentDate.Date Then
            Return DownloadAsk(symbol, targetDate)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows()
                If myRow("Ticker").trim() = symbol And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Ask"))
                End If
            Next
        Else
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTable").Rows()
                If myRow("Symbol").trim() = symbol And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Ask"))
                End If
            Next
        End If

        MessageBox.Show("Holy Batcape! Could not find the ask for " + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetBid(symbol As String, targetDate As Date) As Double
        If targetDate.Date <> currentDate.Date Then
            Return DownloadBid(symbol, targetDate)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows()
                If myRow("Ticker").trim() = symbol And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Bid"))
                End If
            Next
        Else
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTable").Rows()
                If myRow("Symbol").trim() = symbol And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Bid"))
                End If
            Next
        End If

        MessageBox.Show("Holy Batcape! Could not find the bid for " + symbol + ". Returned 0.")
        Return 0
    End Function

End Module
