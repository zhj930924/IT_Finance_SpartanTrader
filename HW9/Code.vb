
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

Module DBprocedures
    ' These are the ADO components
    Dim myConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection
    Dim myConnectionString As String = ""
    Dim myCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter
    Public myDataSet As DataSet = New DataSet

    Public Sub SetUpTheADOcomponents()
        'give to the command a connection
        myCommand.Connection = myConnection
        'give to the data adapter a command
        myDataAdapter.SelectCommand = myCommand
    End Sub

    Public Sub ConnectToDB(connString As String)
        'set the connection string
        myConnection.ConnectionString = connString
        myConnection.Open()
    End Sub

    Public Sub DisconnectFromDB()
        myConnection.Close()
    End Sub

    Public Sub RunQueryAndSaveResultInDS(query As String, resultName As String)
        myCommand.CommandText = query
        myDataAdapter.Fill(myDataSet, resultName)
    End Sub

    Public Sub ClearTableInDS(tableName As String)
        If myDataSet.Tables.Contains(tableName) Then
            myDataSet.Tables(tableName).Clear()
        End If
    End Sub
End Module
