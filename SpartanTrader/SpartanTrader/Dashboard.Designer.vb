﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On



'''
<Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(1),  _
 Global.System.Security.Permissions.PermissionSetAttribute(Global.System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>  _
Partial Public NotInheritable Class Dashboard
    Inherits Microsoft.Office.Tools.Excel.WorksheetBase
    
    Friend WithEvents CAccountCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents MarginCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents MarginPercCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents maxMarginCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TPVatStartCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents IPCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents APCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TPVCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TaTPVCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TECell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TEPercCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents CurrentDateCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents InterestSLTCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TrackingChartLO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents TrackingChart As Microsoft.Office.Tools.Excel.Chart
    
    Friend WithEvents TypeCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents SymbolCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents StrikeCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents DeltaCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents QtyCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents PriceCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TransCostCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TotValueCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents CAccountATCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents MarginATCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TickerCBox As Microsoft.Office.Tools.Excel.Controls.ComboBox
    
    Friend WithEvents StockQtyTBox As Microsoft.Office.Tools.Excel.Controls.TextBox
    
    Friend WithEvents BuyStockBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellStockBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellShortStockBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents CashDivBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExecuteStockTransactionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SymbolCBox As Microsoft.Office.Tools.Excel.Controls.ComboBox
    
    Friend WithEvents OptionQtyTBox As Microsoft.Office.Tools.Excel.Controls.TextBox
    
    Friend WithEvents BuyOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellShortOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExerciseOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExecuteOptionTransactionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Public Sub New(ByVal factory As Global.Microsoft.Office.Tools.Excel.Factory, ByVal serviceProvider As Global.System.IServiceProvider)
        MyBase.New(factory, serviceProvider, "Sheet1", "Sheet1")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub Initialize()
        MyBase.Initialize
        Globals.Dashboard = Me
        Global.System.Windows.Forms.Application.EnableVisualStyles
        Me.InitializeCachedData
        Me.InitializeControls
        Me.InitializeComponents
        Me.InitializeData
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub FinishInitialization()
        Me.OnStartup
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub InitializeDataBindings()
        Me.BeginInitialization
        Me.BindToData
        Me.EndInitialization
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeCachedData()
        If (Me.DataHost Is Nothing) Then
            Return
        End If
        If Me.DataHost.IsCacheInitialized Then
            Me.DataHost.FillCachedData(Me)
        End If
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeData()
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub BindToData()
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Sub StartCaching(ByVal MemberName As String)
        Me.DataHost.StartCaching(Me, MemberName)
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Sub StopCaching(ByVal MemberName As String)
        Me.DataHost.StopCaching(Me, MemberName)
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Function IsCached(ByVal MemberName As String) As Boolean
        Return Me.DataHost.IsCached(Me, MemberName)
    End Function
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub BeginInitialization()
        Me.BeginInit
        Me.CAccountCell.BeginInit
        Me.MarginCell.BeginInit
        Me.MarginPercCell.BeginInit
        Me.maxMarginCell.BeginInit
        Me.TPVatStartCell.BeginInit
        Me.IPCell.BeginInit
        Me.APCell.BeginInit
        Me.TPVCell.BeginInit
        Me.TaTPVCell.BeginInit
        Me.TECell.BeginInit
        Me.TEPercCell.BeginInit
        Me.CurrentDateCell.BeginInit
        Me.InterestSLTCell.BeginInit
        Me.TrackingChartLO.BeginInit
        Me.TrackingChart.BeginInit
        Me.TypeCell.BeginInit
        Me.SymbolCell.BeginInit
        Me.StrikeCell.BeginInit
        Me.DeltaCell.BeginInit
        Me.QtyCell.BeginInit
        Me.PriceCell.BeginInit
        Me.TransCostCell.BeginInit
        Me.TotValueCell.BeginInit
        Me.CAccountATCell.BeginInit
        Me.MarginATCell.BeginInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub EndInitialization()
        Me.MarginATCell.EndInit
        Me.CAccountATCell.EndInit
        Me.TotValueCell.EndInit
        Me.TransCostCell.EndInit
        Me.PriceCell.EndInit
        Me.QtyCell.EndInit
        Me.DeltaCell.EndInit
        Me.StrikeCell.EndInit
        Me.SymbolCell.EndInit
        Me.TypeCell.EndInit
        Me.TrackingChart.EndInit
        Me.TrackingChartLO.EndInit
        Me.InterestSLTCell.EndInit
        Me.CurrentDateCell.EndInit
        Me.TEPercCell.EndInit
        Me.TECell.EndInit
        Me.TaTPVCell.EndInit
        Me.TPVCell.EndInit
        Me.APCell.EndInit
        Me.IPCell.EndInit
        Me.TPVatStartCell.EndInit
        Me.maxMarginCell.EndInit
        Me.MarginPercCell.EndInit
        Me.MarginCell.EndInit
        Me.CAccountCell.EndInit
        Me.EndInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeControls()
        Me.CAccountCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "CAccountCell", "CAccountCell", Me)
        Me.MarginCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "MarginCell", "MarginCell", Me)
        Me.MarginPercCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "MarginPercCell", "MarginPercCell", Me)
        Me.maxMarginCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "maxMarginCell", "maxMarginCell", Me)
        Me.TPVatStartCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TPVatStartCell", "TPVatStartCell", Me)
        Me.IPCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "IPCell", "IPCell", Me)
        Me.APCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "APCell", "APCell", Me)
        Me.TPVCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TPVCell", "TPVCell", Me)
        Me.TaTPVCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TaTPVCell", "TaTPVCell", Me)
        Me.TECell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TECell", "TECell", Me)
        Me.TEPercCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TEPercCell", "TEPercCell", Me)
        Me.CurrentDateCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "CurrentDateCell", "CurrentDateCell", Me)
        Me.InterestSLTCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "InterestSLTCell", "InterestSLTCell", Me)
        Me.TrackingChartLO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet1:TrackingChartLO", "TrackingChartLO", Me)
        Me.TrackingChart = Globals.Factory.CreateChart(Nothing, Nothing, "Sheet1:Chart 3", "TrackingChart", Me)
        Me.TypeCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TypeCell", "TypeCell", Me)
        Me.SymbolCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "SymbolCell", "SymbolCell", Me)
        Me.StrikeCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "StrikeCell", "StrikeCell", Me)
        Me.DeltaCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "DeltaCell", "DeltaCell", Me)
        Me.QtyCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "QtyCell", "QtyCell", Me)
        Me.PriceCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "PriceCell", "PriceCell", Me)
        Me.TransCostCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TransCostCell", "TransCostCell", Me)
        Me.TotValueCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TotValueCell", "TotValueCell", Me)
        Me.CAccountATCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "CAccountATCell", "CAccountATCell", Me)
        Me.MarginATCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "MarginATCell", "MarginATCell", Me)
        Me.TickerCBox = New Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "1DD05228C1726B1414D1B95A126299A2B52431", "1DD05228C1726B1414D1B95A126299A2B52431", Me, "TickerCBox")
        Me.StockQtyTBox = New Microsoft.Office.Tools.Excel.Controls.TextBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "25E36214728760248CC2B2BC260D5B57C6DE52", "25E36214728760248CC2B2BC260D5B57C6DE52", Me, "StockQtyTBox")
        Me.BuyStockBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "3A481986333AC4343973BE5F30EC0D577E9343", "3A481986333AC4343973BE5F30EC0D577E9343", Me, "BuyStockBtn")
        Me.SellStockBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "4396E5BDA4AFAE44C254AB7D48395781A553B4", "4396E5BDA4AFAE44C254AB7D48395781A553B4", Me, "SellStockBtn")
        Me.SellShortStockBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "583C43A6C55DE154C655BC2A5144E93A287F75", "583C43A6C55DE154C655BC2A5144E93A287F75", Me, "SellShortStockBtn")
        Me.CashDivBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "63B5D4A9A6B283641286A7BA61C6BF7C99D8C6", "63B5D4A9A6B283641286A7BA61C6BF7C99D8C6", Me, "CashDivBtn")
        Me.ExecuteStockTransactionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "7460F52E77EE43749957AA857E1EFDF6AE0337", "7460F52E77EE43749957AA857E1EFDF6AE0337", Me, "ExecuteStockTransactionBtn")
        Me.SymbolCBox = New Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "8840B188F8E99284492884D982FE494AC5F828", "8840B188F8E99284492884D982FE494AC5F828", Me, "SymbolCBox")
        Me.OptionQtyTBox = New Microsoft.Office.Tools.Excel.Controls.TextBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "9E7C1056F984539436D980C998BCFF3C4D3659", "9E7C1056F984539436D980C998BCFF3C4D3659", Me, "OptionQtyTBox")
        Me.BuyOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "12470469A1D9831473819CFA1F9FEEA176F131", "12470469A1D9831473819CFA1F9FEEA176F131", Me, "BuyOptionBtn")
        Me.SellOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "13E17871A18C1B14C5B1AE6C195D8AF5E1FFB1", "13E17871A18C1B14C5B1AE6C195D8AF5E1FFB1", Me, "SellOptionBtn")
        Me.SellShortOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1BAE784C1191B314A841BF0713EAAEDF700291", "1BAE784C1191B314A841BF0713EAAEDF700291", Me, "SellShortOptionBtn")
        Me.ExerciseOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1C747B3C41B855145791A266199CDFFB7F29F1", "1C747B3C41B855145791A266199CDFFB7F29F1", Me, "ExerciseOptionBtn")
        Me.ExecuteOptionTransactionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1A36E87D316A7E149A01820A11C732994BE281", "1A36E87D316A7E149A01820A11C732994BE281", Me, "ExecuteOptionTransactionBtn")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeComponents()
        '
        'TickerCBox
        '
        Me.TickerCBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.TickerCBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.TickerCBox.ForeColor = System.Drawing.Color.Transparent
        Me.TickerCBox.Name = "TickerCBox"
        Me.TickerCBox.Text = "Select Ticker"
        '
        'StockQtyTBox
        '
        Me.StockQtyTBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.StockQtyTBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.StockQtyTBox.ForeColor = System.Drawing.Color.Transparent
        Me.StockQtyTBox.Name = "StockQtyTBox"
        Me.StockQtyTBox.Text = "0"
        '
        'BuyStockBtn
        '
        Me.BuyStockBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.BuyStockBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.BuyStockBtn.ForeColor = System.Drawing.Color.Transparent
        Me.BuyStockBtn.Name = "BuyStockBtn"
        Me.BuyStockBtn.Text = "Buy"
        Me.BuyStockBtn.UseVisualStyleBackColor = false
        '
        'SellStockBtn
        '
        Me.SellStockBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.SellStockBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.SellStockBtn.ForeColor = System.Drawing.Color.Transparent
        Me.SellStockBtn.Name = "SellStockBtn"
        Me.SellStockBtn.Text = "Sell"
        Me.SellStockBtn.UseVisualStyleBackColor = false
        '
        'SellShortStockBtn
        '
        Me.SellShortStockBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.SellShortStockBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.SellShortStockBtn.ForeColor = System.Drawing.Color.Transparent
        Me.SellShortStockBtn.Name = "SellShortStockBtn"
        Me.SellShortStockBtn.Text = "SellShort"
        Me.SellShortStockBtn.UseVisualStyleBackColor = false
        '
        'CashDivBtn
        '
        Me.CashDivBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.CashDivBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.CashDivBtn.ForeColor = System.Drawing.Color.Transparent
        Me.CashDivBtn.Name = "CashDivBtn"
        Me.CashDivBtn.Text = "Cash Div"
        Me.CashDivBtn.UseVisualStyleBackColor = false
        '
        'ExecuteStockTransactionBtn
        '
        Me.ExecuteStockTransactionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.ExecuteStockTransactionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.ExecuteStockTransactionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.ExecuteStockTransactionBtn.Name = "ExecuteStockTransactionBtn"
        Me.ExecuteStockTransactionBtn.Text = "Execute Stock Transaction"
        Me.ExecuteStockTransactionBtn.UseVisualStyleBackColor = false
        '
        'SymbolCBox
        '
        Me.SymbolCBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.SymbolCBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.SymbolCBox.ForeColor = System.Drawing.Color.Transparent
        Me.SymbolCBox.Name = "SymbolCBox"
        Me.SymbolCBox.Text = "Select Symbol"
        '
        'OptionQtyTBox
        '
        Me.OptionQtyTBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.OptionQtyTBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.OptionQtyTBox.ForeColor = System.Drawing.Color.Transparent
        Me.OptionQtyTBox.Name = "OptionQtyTBox"
        Me.OptionQtyTBox.Text = "0"
        '
        'BuyOptionBtn
        '
        Me.BuyOptionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.BuyOptionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.BuyOptionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.BuyOptionBtn.Name = "BuyOptionBtn"
        Me.BuyOptionBtn.Text = "Buy"
        Me.BuyOptionBtn.UseVisualStyleBackColor = false
        '
        'SellOptionBtn
        '
        Me.SellOptionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.SellOptionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.SellOptionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.SellOptionBtn.Name = "SellOptionBtn"
        Me.SellOptionBtn.Text = "Sell"
        Me.SellOptionBtn.UseVisualStyleBackColor = false
        '
        'SellShortOptionBtn
        '
        Me.SellShortOptionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.SellShortOptionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.SellShortOptionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.SellShortOptionBtn.Name = "SellShortOptionBtn"
        Me.SellShortOptionBtn.Text = "SellShort"
        Me.SellShortOptionBtn.UseVisualStyleBackColor = false
        '
        'ExerciseOptionBtn
        '
        Me.ExerciseOptionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.ExerciseOptionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.ExerciseOptionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.ExerciseOptionBtn.Name = "ExerciseOptionBtn"
        Me.ExerciseOptionBtn.Text = "Exercise"
        Me.ExerciseOptionBtn.UseVisualStyleBackColor = false
        '
        'ExecuteOptionTransactionBtn
        '
        Me.ExecuteOptionTransactionBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(32,Byte),Integer), CType(CType(56,Byte),Integer), CType(CType(100,Byte),Integer))
        Me.ExecuteOptionTransactionBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.ExecuteOptionTransactionBtn.ForeColor = System.Drawing.Color.Transparent
        Me.ExecuteOptionTransactionBtn.Name = "ExecuteOptionTransactionBtn"
        Me.ExecuteOptionTransactionBtn.Text = "Execute Option Transaction"
        Me.ExecuteOptionTransactionBtn.UseVisualStyleBackColor = false
        '
        'CAccountCell
        '
        Me.CAccountCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'MarginCell
        '
        Me.MarginCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'MarginPercCell
        '
        Me.MarginPercCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'maxMarginCell
        '
        Me.maxMarginCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TPVatStartCell
        '
        Me.TPVatStartCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'IPCell
        '
        Me.IPCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'APCell
        '
        Me.APCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TPVCell
        '
        Me.TPVCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TaTPVCell
        '
        Me.TaTPVCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TECell
        '
        Me.TECell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TEPercCell
        '
        Me.TEPercCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'CurrentDateCell
        '
        Me.CurrentDateCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'InterestSLTCell
        '
        Me.InterestSLTCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TrackingChartLO
        '
        Me.TrackingChartLO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TrackingChart
        '
        Me.TrackingChart.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TypeCell
        '
        Me.TypeCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'SymbolCell
        '
        Me.SymbolCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'StrikeCell
        '
        Me.StrikeCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'DeltaCell
        '
        Me.DeltaCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'QtyCell
        '
        Me.QtyCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'PriceCell
        '
        Me.PriceCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TransCostCell
        '
        Me.TransCostCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TotValueCell
        '
        Me.TotValueCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'CAccountATCell
        '
        Me.CAccountATCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'MarginATCell
        '
        Me.MarginATCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'Dashboard
        '
        Me.TickerCBox.BindingContext = Me.BindingContext
        Me.StockQtyTBox.BindingContext = Me.BindingContext
        Me.BuyStockBtn.BindingContext = Me.BindingContext
        Me.SellStockBtn.BindingContext = Me.BindingContext
        Me.SellShortStockBtn.BindingContext = Me.BindingContext
        Me.CashDivBtn.BindingContext = Me.BindingContext
        Me.ExecuteStockTransactionBtn.BindingContext = Me.BindingContext
        Me.SymbolCBox.BindingContext = Me.BindingContext
        Me.OptionQtyTBox.BindingContext = Me.BindingContext
        Me.BuyOptionBtn.BindingContext = Me.BindingContext
        Me.SellOptionBtn.BindingContext = Me.BindingContext
        Me.SellShortOptionBtn.BindingContext = Me.BindingContext
        Me.ExerciseOptionBtn.BindingContext = Me.BindingContext
        Me.ExecuteOptionTransactionBtn.BindingContext = Me.BindingContext
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Function NeedsFill(ByVal MemberName As String) As Boolean
        Return Me.DataHost.NeedsFill(Me, MemberName)
    End Function
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub OnShutdown()
        Me.MarginATCell.Dispose
        Me.CAccountATCell.Dispose
        Me.TotValueCell.Dispose
        Me.TransCostCell.Dispose
        Me.PriceCell.Dispose
        Me.QtyCell.Dispose
        Me.DeltaCell.Dispose
        Me.StrikeCell.Dispose
        Me.SymbolCell.Dispose
        Me.TypeCell.Dispose
        Me.TrackingChart.Dispose
        Me.TrackingChartLO.Dispose
        Me.InterestSLTCell.Dispose
        Me.CurrentDateCell.Dispose
        Me.TEPercCell.Dispose
        Me.TECell.Dispose
        Me.TaTPVCell.Dispose
        Me.TPVCell.Dispose
        Me.APCell.Dispose
        Me.IPCell.Dispose
        Me.TPVatStartCell.Dispose
        Me.maxMarginCell.Dispose
        Me.MarginPercCell.Dispose
        Me.MarginCell.Dispose
        Me.CAccountCell.Dispose
        MyBase.OnShutdown
    End Sub
End Class

Partial Friend NotInheritable Class Globals
    
    Private Shared _Dashboard As Dashboard
    
    Friend Shared Property Dashboard() As Dashboard
        Get
            Return _Dashboard
        End Get
        Set
            If (_Dashboard Is Nothing) Then
                _Dashboard = value
            Else
                Throw New System.NotSupportedException()
            End If
        End Set
    End Property
End Class
