
Public Class Sheet1

    Private Sub Sheet1_Startup() Handles Me.Startup
        BudgetDataLst.AutoSetDataBoundColumnHeaders = True
        SetUpTheADOcomponents()
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown
        DisconnectFromDB()
    End Sub

    Private Sub LoadBudgetDataBtn_Click(sender As Object, e As EventArgs) Handles LoadBudgetDataBtn.Click
        ClearTableInDS("BudgetDataTbl")
        RunQueryAndSaveResultInDS("Select * from Acme_Budget", "BudgetDataTbl")
        BudgetDataLst.DataSource = myDataSet.Tables("BudgetDataTbl")
    End Sub

    Private Sub FormatAsAPivotBtn_Click(sender As Object, e As EventArgs) Handles FormatAsAPivotBtn.Click
        'Declare the pivot objects
        Dim myPivotCache As Excel.PivotCache
        Dim myPivotTable As Excel.PivotTable
        Dim myDataSheet As Excel.Worksheet = Application.ActiveWorkbook.ActiveSheet
        Dim myNewSheet As Excel.Worksheet = Application.ActiveWorkbook.Worksheets.Add()

        'Find the last filled row
        Dim lastRow = myDataSheet.Cells(Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

        'Create the PivotCache from data in the data sheet
        myPivotCache = Application.ActiveWorkbook.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase,
                                                                     myDataSheet.Range("A1:G" & lastRow))
        'Create the PivotTable
        myPivotTable = myPivotCache.CreatePivotTable(myNewSheet.Range("A1"), "myPivotCache")

        'Add fields to Report Filter
        myPivotTable.PivotFields("Item").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        myPivotTable.PivotFields("Category").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        myPivotTable.PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        'Add fields to Column Labels
        myPivotTable.PivotFields("Month").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'Add fields to Row Labels
        myPivotTable.PivotFields("Department").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        'Add fields to Values
        myPivotTable.PivotFields("Actual").Orientation = Excel.XlPivotFieldOrientation.xlDataField

        'Add calculated field
        myPivotTable.CalculatedFields.Add("Variance", "=Budget-Actual")
        myPivotTable.PivotFields("Variance").Orientation = Excel.XlPivotFieldOrientation.xlDataField

        'Format Numbers
        myPivotTable.DataBodyRange.NumberFormat = "$ #,##0;[Red]($ #,##0)"

        'Style
        myPivotTable.TableStyle2 = "PivotStyleMedium2"
        myPivotTable.DisplayFieldCaptions = False

        'Change the captions
        myPivotTable.PivotFields("Sum of Actual").Caption = " Actual"
        myPivotTable.PivotFields("Sum of Variance").Caption = " Variance"

        'Set the orientation of the sums
        myPivotTable.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField

    End Sub
End Class
