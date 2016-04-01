Module PortfolioManagement

    Public Function GetCurrPositionInAP(symbol) As Integer

        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            If myRow("Symbol").ToString().Trim() = symbol Then
                Return Integer.Parse(myRow("Units"))
            End If
        Next
        Return 0

    End Function

    '-- Homework 16---------------------------------------------------------------------------------------------
    Public Function CalcTPVNoHedge(targetdate As Date) As Double

        Dim ts As TimeSpan = targetdate.Date - lastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        Dim interest As Double = CAccount * (Math.Exp(iRate * t) - 1)
        Return (CalcIPValue(targetdate) + initialCAccount + interest)

    End Function

    Public Function CalcTPV(targetdate As Date) As Double

        Return (CalcIPValue(targetdate) + CalcAPValue(targetdate) + CAccount + CalcInterestSLT(targetdate))

    End Function

    Public Function CalcMargin(targetdate As Date) As Double

        Dim tempMargin As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Integer

        ' First, margins for IP
        If myDataSet.Tables.Contains("InitialPositionTable") Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                tempSymbol = myRow("Symbol").ToString().Trim
                tempUnits = myRow("Units")
                If tempUnits < 0 Then ' add if position is short
                    tempMargin = tempMargin + (-tempUnits * CalcMTM(tempSymbol, targetdate))
                    ' the minus sign makes the units positive
                End If
            Next
        End If

        ' Next, margin for AP
        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            tempSymbol = myRow("Symbol").ToString().Trim
            tempUnits = myRow("Units")
            If (tempUnits < 0) And (tempSymbol <> "CAccount") Then ' add if position is short
                ' add if position is short and it is not the CAccount
                tempMargin = tempMargin + (-tempUnits * CalcMTM(tempSymbol, targetdate))
                ' the minus sign makes the units positive
            End If
        Next
        Return tempMargin

    End Function

    Public Function CalcAPValue(targetdate As Date) As Double

        Dim tempAP As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double

        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            tempSymbol = myRow("Symbol").ToString().Trim
            tempUnits = myRow("Units")
            If tempSymbol <> "CAccount" Then
                tempAP = tempAP + (tempUnits * CalcMTM(tempSymbol, targetdate))
            End If
        Next
        Return tempAP

    End Function

    Public Function CalcInterestSLT(toThisDay As Date) As Double

        Dim interest As Double = 0
        Dim ts As TimeSpan = toThisDay.Date - lastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        interest = CAccount * (Math.Exp(iRate * t) - 1)
        Return interest

    End Function

    Public Function CalcTaTPV(targetdate As Date) As Double

        Dim ts As TimeSpan = targetdate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25
        Return TPVatStart * Math.Exp(iRate * t)

    End Function




    ' -----------------------------homework 14------------------------------------------------------------
    Public Function CalcTPVAtStart() As Double

        Return CalcIPValue(startDate) + initialCAccount

    End Function

    Public Function CalcIPValue(targetDate As Date) As Double

        Dim tempCumulativeValue As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double

        If myDataSet.Tables.Contains("InitialPositionTable") Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                tempSymbol = myRow("Symbol").ToString().Trim
                tempUnits = myRow("Units")
                tempCumulativeValue = tempCumulativeValue + (tempUnits * CalcMTM(tempSymbol, targetDate))
            Next
        End If

        Return tempCumulativeValue

    End Function

    Public Function CalcMTM(symbol As String, targetDate As Date) As Double

        Return (GetAsk(symbol, targetDate) + GetBid(symbol, targetDate)) / 2

    End Function

    Public Function IsAStock(Symbol As String) As Boolean

        Symbol = Symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("TickerTable").Rows
            If myRow("Ticker").trim = Symbol Then
                Return True
            End If
        Next
        Return False

    End Function

End Module
