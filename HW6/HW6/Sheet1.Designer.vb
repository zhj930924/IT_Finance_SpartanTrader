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
    
    Friend WithEvents StartBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents CleanPhNoBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents FormatDatesBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ComputeCagrBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
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
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub EndInitialization()
        Me.EndInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeControls()
        Me.StartBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "29B7825252AC2D24DB2284BC203140F4ECCCF2", "29B7825252AC2D24DB2284BC203140F4ECCCF2", Me, "StartBtn")
        Me.CleanPhNoBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "30A08DA9C32D67348C03838F3E1B55FF4D6773", "30A08DA9C32D67348C03838F3E1B55FF4D6773", Me, "CleanPhNoBtn")
        Me.FormatDatesBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "41FC1CB704EFD74422E4BCFE41FB70ED3FF844", "41FC1CB704EFD74422E4BCFE41FB70ED3FF844", Me, "FormatDatesBtn")
        Me.ComputeCagrBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "5BC7508AF55116546915B5185F5CDFCEF7FEF5", "5BC7508AF55116546915B5185F5CDFCEF7FEF5", Me, "ComputeCagrBtn")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeComponents()
        '
        'StartBtn
        '
        Me.StartBtn.BackColor = System.Drawing.Color.Turquoise
        Me.StartBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.StartBtn.ForeColor = System.Drawing.Color.Black
        Me.StartBtn.Name = "StartBtn"
        Me.StartBtn.Text = "Start"
        Me.StartBtn.UseVisualStyleBackColor = false
        '
        'CleanPhNoBtn
        '
        Me.CleanPhNoBtn.BackColor = System.Drawing.Color.PaleTurquoise
        Me.CleanPhNoBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold)
        Me.CleanPhNoBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CleanPhNoBtn.Name = "CleanPhNoBtn"
        Me.CleanPhNoBtn.Text = "Clean Phone Numbers"
        Me.CleanPhNoBtn.UseVisualStyleBackColor = false
        '
        'FormatDatesBtn
        '
        Me.FormatDatesBtn.BackColor = System.Drawing.Color.LightSeaGreen
        Me.FormatDatesBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold)
        Me.FormatDatesBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormatDatesBtn.Name = "FormatDatesBtn"
        Me.FormatDatesBtn.Text = "Format Dates"
        Me.FormatDatesBtn.UseVisualStyleBackColor = false
        '
        'ComputeCagrBtn
        '
        Me.ComputeCagrBtn.BackColor = System.Drawing.Color.MediumTurquoise
        Me.ComputeCagrBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold)
        Me.ComputeCagrBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ComputeCagrBtn.Name = "ComputeCagrBtn"
        Me.ComputeCagrBtn.Text = "Compute CAGR"
        Me.ComputeCagrBtn.UseVisualStyleBackColor = false
        '
        'Sheet1
        '
        Me.StartBtn.BindingContext = Me.BindingContext
        Me.CleanPhNoBtn.BindingContext = Me.BindingContext
        Me.FormatDatesBtn.BindingContext = Me.BindingContext
        Me.ComputeCagrBtn.BindingContext = Me.BindingContext
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Function NeedsFill(ByVal MemberName As String) As Boolean
        Return Me.DataHost.NeedsFill(Me, MemberName)
    End Function
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