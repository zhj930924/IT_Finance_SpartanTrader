Imports Microsoft.Office.Tools.Ribbon

Public Class stRibbon

    Private Sub TransactionQueueBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionQueueBtn.Click
        DownloadTransactionQueue(teamID)
        ShowTransactionQueue()
    End Sub

    Private Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click
        Globals.Dashboard.Activate()
    End Sub

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
