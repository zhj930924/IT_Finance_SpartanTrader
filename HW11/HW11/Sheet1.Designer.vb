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
Partial Public NotInheritable Class Sheet1
    Inherits Microsoft.Office.Tools.Excel.WorksheetBase
    
    Friend WithEvents StockDataLst As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents LatestClosingsBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents HistForTickerBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents TickerCBox As Microsoft.Office.Tools.Excel.Controls.ComboBox
    
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
        Globals.Sheet1 = Me
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
        Me.StockDataLst.BeginInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub EndInitialization()
        Me.StockDataLst.EndInit
        Me.EndInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeControls()
        Me.StockDataLst = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet1:StockDataLst", "StockDataLst", Me)
        Me.LatestClosingsBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1AE7462391ACC814B2F1B7A31AF019E4958F11", "1AE7462391ACC814B2F1B7A31AF019E4958F11", Me, "LatestClosingsBtn")
        Me.HistForTickerBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "235D6B86C2C2A6242E52B7C627895C030EB252", "235D6B86C2C2A6242E52B7C627895C030EB252", Me, "HistForTickerBtn")
        Me.TickerCBox = New Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "3116478DD3D2C534D583B99432CEFB70658393", "3116478DD3D2C534D583B99432CEFB70658393", Me, "TickerCBox")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeComponents()
        '
        'LatestClosingsBtn
        '
        Me.LatestClosingsBtn.BackColor = System.Drawing.Color.Black
        Me.LatestClosingsBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.LatestClosingsBtn.ForeColor = System.Drawing.Color.Transparent
        Me.LatestClosingsBtn.Name = "LatestClosingsBtn"
        Me.LatestClosingsBtn.Text = "Lastest Closings"
        Me.LatestClosingsBtn.UseVisualStyleBackColor = false
        '
        'HistForTickerBtn
        '
        Me.HistForTickerBtn.BackColor = System.Drawing.Color.SteelBlue
        Me.HistForTickerBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold)
        Me.HistForTickerBtn.ForeColor = System.Drawing.Color.Transparent
        Me.HistForTickerBtn.Name = "HistForTickerBtn"
        Me.HistForTickerBtn.Text = "Historic For Ticker"
        Me.HistForTickerBtn.UseVisualStyleBackColor = false
        '
        'TickerCBox
        '
        Me.TickerCBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.TickerCBox.Name = "TickerCBox"
        Me.TickerCBox.Text = "Select a ticker"
        '
        'StockDataLst
        '
        Me.StockDataLst.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'Sheet1
        '
        Me.LatestClosingsBtn.BindingContext = Me.BindingContext
        Me.HistForTickerBtn.BindingContext = Me.BindingContext
        Me.TickerCBox.BindingContext = Me.BindingContext
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
        Me.StockDataLst.Dispose
        MyBase.OnShutdown
    End Sub
End Class

Partial Friend NotInheritable Class Globals
    
    Private Shared _Sheet1 As Sheet1
    
    Friend Shared Property Sheet1() As Sheet1
        Get
            Return _Sheet1
        End Get
        Set
            If (_Sheet1 Is Nothing) Then
                _Sheet1 = value
            Else
                Throw New System.NotSupportedException()
            End If
        End Set
    End Property
End Class
