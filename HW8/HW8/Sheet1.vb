
Public Class Sheet1

    Private Sub Sheet1_Startup() Handles Me.Startup
        CustomersLst.AutoSetDataBoundColumnHeaders = True
        SetUpTheADOcomponents()
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub AllCustomersBtn_Click(sender As Object, e As EventArgs) Handles AllCustomersBtn.Click
        ClearTableInDS("CustomersTbl")
        ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;Initial Catalog=SmallBankDB;Integrated Security=True")
        RunQueryAndSaveResultInDS("Select * From Customer", "Customer")
        CustomersLst.DataSource = myDataSet.Tables("CustomerTbl")
        DisconnectFromDB()
    End Sub

End Class
