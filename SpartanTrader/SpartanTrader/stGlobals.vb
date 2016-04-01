Module stGlobals

    Public CAccountAT As Double = 0
    Public marginAT As Double = 0

    ' ---------- homework 16 --------------------------------------

    Public myTransaction As Transaction = New Transaction

    ' ---------- homework 15 --------------------------------------

    Public CAccount As Double = 0
    Public margin As Double = 0
    Public AP As Double = 0
    Public TPV As Double = 0
    Public TaTPV As Double = 0
    Public TE As Double = 0
    Public lastTransactionDate As Date
    Public interestSLT As Double = 0
    Public TEpercent As Double = 0
    Public lastPriceDownloadDate As Date

    ' ---------- homework 14 --------------------------------------

    Public initialCAccount As Double = 0
    Public iRate As Double = 0
    Public startDate As Date
    Public currentDate As Date
    Public endDate As Date
    Public maxMargins As Double = 0
    Public TPVatStart As Double = 0
    Public IP As Double = 0

    ' ---------- homework 13 --------------------------------------
    Public activeDB As String = ""
    Public teamID As String = "30"
    Public portfolioTableName As String = "PortfolioTeam" + teamID

End Module
