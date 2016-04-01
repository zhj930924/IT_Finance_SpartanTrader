
Public Class Sheet1
    Private Sub Sheet1_Startup() Handles Me.Startup
        CustomersLst.AutoSetDataBoundColumnHeaders = True
        SetUpTheADOcomponents()
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub
    Private Sub AllCustomersBtn_Click(sender As Object, e As EventArgs) Handles AllCustomersBtn.Click
        ClearTableInDS("CustomerTbl")
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
        RunQueryAndSaveResultInDS("Select * From Customer Order by l_name", "CustomerTbl")
        CustomersLst.DataSource = myDataSet.Tables("CustomerTbl")
        DisconnectFromDB()
    End Sub

    Private Sub GetVaLoanBtn_Click(sender As Object, e As EventArgs) Handles GetVaLoanBtn.Click
        ClearTableInDS("VaLoanTbl")
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
        RunQueryAndSaveResultInDS("SELECT CUSTOMER.c_id, f_name, l_name, LOAN.* FROM CUSTOMER, LOAN, CUSTOMER_IN_LOAN
                                    WHERE CUSTOMER.c_id = CUSTOMER_IN_LOAN.c_id
                                    AND LOAN.l_id = CUSTOMER_IN_LOAN.l_id
                                    AND state = 'VA' Order By l_name", "VaLoanTbl")
        CustomersLst.DataSource = myDataSet.Tables("VaLoanTbl")
        DisconnectFromDB()
    End Sub

    Private Sub GetTxLoanBtn_Click(sender As Object, e As EventArgs) Handles GetTxLoanBtn.Click
        ClearTableInDS("TxLoanTbl")
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
        RunQueryAndSaveResultInDS("SELECT CUSTOMER.c_id, f_name, l_name, LOAN.* FROM CUSTOMER, LOAN, CUSTOMER_IN_LOAN
                                    WHERE CUSTOMER.c_id = CUSTOMER_IN_LOAN.c_id
                                    AND LOAN.l_id = CUSTOMER_IN_LOAN.l_id
                                    AND state = 'TX' Order By l_name", "TxLoanTbl")
        CustomersLst.DataSource = myDataSet.Tables("TxLoanTbl")
        DisconnectFromDB()
    End Sub
End Class
