Public Class Transaction

    Public trType As String = ""

    Public price As Double = 0
    Public action As String = ""
    Public symbol As String = ""
    Public typeOfSecurity As String = ""
    Public qty As Double = 0
    Public transCost As Double = 0
    Public totValue As Double = 0
    Public typeOfPrice As String = "" 'Bid, Ask
    Public optionType As String = "" 'Call or Put
    Public delta As Double = 0
    Public strike As Double = 0

    Public Sub ExecuteTransaction()

        Dim mySQL As String
        mySQL = String.Format("INSERT INTO TransactionQueue (Date, TeamID, Symbol, Type, Qty, Price, Cost, TotValue, " +
                              "InterestSinceLastTransaction, CashPositionAfterTransaction, TotMargin) VALUES " +
                              "('{0}', {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10})",
                              currentDate.ToShortDateString,
                              teamID,
                              symbol,
                              trType,
                              qty,
                              price,
                              transCost,
                              totValue,
                              interestSLT,
                              CAccountAT,
                              marginAT)
        ExecuteNonQuery(mySQL)
        lastTransactionDate = currentDate
        CAccount = CAccountAT
        margin = marginAT
        'Globals.Portfolio.UpdateAP()

    End Sub

    Public Sub ComputeStockTransactionProperties()

        Select Case typeOfPrice
            Case "Bid"
                price = GetBid(symbol, currentDate)
            Case "Ask"
                price = GetAsk(symbol, currentDate)
            Case "Div"
                price = GetDividend(symbol, currentDate)
            Case Else
                price = 0
        End Select

        transCost = CalcTransCost()
        totValue = CalcTotValue()
        interestSLT = CalcInterestSLT(currentDate)
        CAccountAT = CAccount + totValue + interestSLT
        marginAT = margin + EffectOfTransactionOnMargin()

    End Sub

    Public Sub DisplayTransactionData()

        Try
            Globals.Dashboard.PriceCell.Value = price
            Globals.Dashboard.TypeCell.Value = trType
            Globals.Dashboard.SymbolCell.Value = symbol
            Globals.Dashboard.QtyCell.Value = qty
            Globals.Dashboard.TransCostCell.Value = transCost
            Globals.Dashboard.TotValueCell.Value = totValue
            Globals.Dashboard.InterestSLTCell.Value = interestSLT
            Globals.Dashboard.DeltaCell.Value = delta
            Globals.Dashboard.CAccountATCell.Value = CAccountAT
            Globals.Dashboard.MarginATCell.Value = marginAT
        Catch
            'skip
        End Try

    End Sub

    Public Sub Clear()

        price = 0
        trType = ""
        symbol = ""
        typeOfSecurity = ""
        qty = 0
        transCost = 0
        totValue = 0
        typeOfPrice = ""
        optionType = ""
        delta = 0
        strike = 0
        Globals.Dashboard.ClearTransactionHighlight()

    End Sub

    Public Function EffectOfTransactionOnMargin() As Double

        Dim currPosition As Integer = 0
        Dim underlierPosition As Integer = 0
        Dim effect As Double = 0

        Select Case trType

            Case "Sell"
                ' Sell has no effect on margin becasue you can only sell what you have long
                Return 0 'effect or transation on margin

            Case "Buy"
                currPosition = GetCurrPositionInAP(symbol)
                If currPosition >= 0 Then
                    Return 0
                Else
                    If qty >= Math.Abs(currPosition) Then
                        Return currPosition * CalcMTM(symbol, currentDate)
                        ' buying eliminates all margin for this symbol
                    Else
                        Return -(qty * CalcMTM(symbol, currentDate))
                        ' buying reduces the margin
                    End If
                End If

            Case "SellShort"
                Return qty * CalcMTM(symbol, currentDate)
                ' Selling short always increases the margin

        End Select

        Return 0

    End Function

    '-- Homework 16-------------------------------------------------------------------
    Public Function IsStockInputValid() As Boolean

        'Check ticker
        If Globals.Dashboard.TickerCBox.SelectedItem = Nothing Then
            MessageBox.Show("Picking stocks is hard, I know. Do your best, Dave.",
                            "No ticker", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            symbol = Globals.Dashboard.TickerCBox.SelectedItem
        End If

        'Check type
        If action = "" Then
            MessageBox.Show("To buy or not to buy, that is the question.",
                            "No transaction type", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        'qty
        Try
            qty = Integer.Parse(Globals.Dashboard.StockQtyTBox.Text)
        Catch
            MessageBox.Show("Quantity, Dave?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True 'if all checks are passed

    End Function

    Public Function CalcTransCost() As Double

        Return GetTCostCoefficient(symbol, action) * Math.Abs(qty) * price

    End Function

    Public Function CalcTotValue() As Double

        Select Case action
            Case "Buy"
                Return -(price * qty) - transCost
            Case "Sell"
                Return (price * qty) - transCost
            Case "SellShort"
                Return (price * qty) - transCost
            Case "CashDiv"
                Return (price * qty) - transCost
            Case "X-Put"
                Return (price * qty) - transCost
            Case "X-Call"
                Return -(price * qty) - transCost
            Case Else
                Return 0
        End Select

    End Function

End Class
