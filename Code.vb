Module DBprocedures
    ' These are the ADO components
    Dim myConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection
    Dim myConnectionString As String = ""
    Dim myCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter
    Public myDataSet As DataSet = New DataSet

    Public Sub SetUpTheADOcomponents()
        Try
            'give to the command a connection
            myCommand.Connection = myConnection
            'give to the data adapter a command
            myDataAdapter.SelectCommand = myCommand
        Catch e As Exception
            MessageBox.Show("Ehi! SetUpTheADOcomponents failed: " + e.Message)
        End Try
    End Sub

    Public Sub ConnectToDB(connString As String)
        Try
            'set the connection string
            myConnection.ConnectionString = connString
            myConnection.Open()
        Catch e As Exception
            MessageBox.Show("Ehi! ConnectToDB failed: " + e.Message)
        End Try
    End Sub

    Public Sub DisconnectFromDB()
        Try
            myConnection.Close()
        Catch e As Exception
            MessageBox.Show("Ehi! DisconnectFromDB failed: " + e.Message)
        End Try
    End Sub

    Public Sub RunQueryAndSaveResultInDS(query As String, resultName As String)
        Try
            myCommand.CommandText = query
            myDataAdapter.Fill(myDataSet, resultName)
        Catch e As Exception
            MessageBox.Show("Ehi! RunQueryAndSaveResultInDS failed: " + e.Message)
        End Try
    End Sub

    Public Sub ClearTableInDS(tableName As String)
        Try
            If myDataSet.Tables.Contains(tableName) Then
                myDataSet.Tables(tableName).Clear()
            End If
        Catch e As Exception
            MessageBox.Show("Ehi! ClearTableInDS failed: " + e.Message)
        End Try
    End Sub

    Public Sub ExecuteNonQuery(query As String)
        Try
            myCommand.CommandText = query
            myCommand.ExecuteNonQuery()
        Catch e As Exception
            MessageBox.Show("Ehi! ExecuteNonQuery failed: " + e.Message)
        End Try
    End Sub

End Module

Public Class Sheet1
    Private Sub Sheet1_Startup() Handles Me.Startup
        CustomersLst.AutoSetDataBoundColumnHeaders = True
        SetUpTheADOcomponents()
        ConnectToDB("Data Source = f-sg6m-s4.comm.virginia.edu; Initial Catalog = SmallBankDB; Integrated Security = True")
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown
        DisconnectFromDB()
    End Sub
    Private Sub LoadCustTblBtn_Click(sender As Object, e As EventArgs) Handles LoadCustTblBtn.Click
        ClearTableInDS("CustomersTbl")
        RunQueryAndSaveResultInDS("SELECT * FROM Customer2", "CustomersTbl")
        CustomersLst.DataSource = myDataSet.Tables("CustomersTbl")
    End Sub

    Private Sub DeleteRowBtn_Click(sender As Object, e As EventArgs) Handles DeleteRowBtn.Click
        Dim selectedRow As Integer = Application.ActiveCell.Row
        Dim cIdToDelete As String = Range("A" & selectedRow).Value
        Dim myString As String = String.Format("DELETE FROM Customer2 WHERE c_id = '{0}'",
                                               cIdToDelete)
        ExecuteNonQuery(myString)
        LoadCustTblBtn_Click(Nothing, Nothing)
    End Sub

    Private Sub UpdateRowBtn_Click(sender As Object, e As EventArgs) Handles UpdateRowBtn.Click
        Dim selectedRow As Integer = Application.ActiveCell.Row
        Dim cIdToUpdate As String = Range("A" & selectedRow).Value
        Dim myString As String = ""
        Dim newValue As String = ""

        If Application.ActiveCell.Cells.Column = 1 Then
            MessageBox.Show("You cannot update the identifier.")
            Return
        End If

        Range("A" & selectedRow & ":E" & selectedRow).Select()

        'f_name
        newValue = Range("B" + selectedRow.ToString).Value.ToString()
        myString = String.Format("UPDATE Customer2 SET f_name = '{0}' WHERE c_id = '{1}'",
                                 newValue,
                                 cIdToUpdate)
        ExecuteNonQuery(myString)

        'l_name
        newValue = Range("C" + selectedRow.ToString).Value.ToString()
        myString = String.Format("UPDATE Customer2 SET l_name = '{0}' WHERE c_id = '{1}'",
                                 newValue,
                                 cIdToUpdate)
        ExecuteNonQuery(myString)

        'city
        newValue = Range("D" + selectedRow.ToString).Value.ToString()
        myString = String.Format("UPDATE Customer2 SET city = '{0}' WHERE c_id = '{1}'",
                                 newValue,
                                 cIdToUpdate)
        ExecuteNonQuery(myString)

        'state
        newValue = Range("E" + selectedRow.ToString).Value.ToString()
        myString = String.Format("UPDATE Customer2 SET state = '{0}' WHERE c_id = '{1}'",
                                 newValue,
                                 cIdToUpdate)
        ExecuteNonQuery(myString)

        LoadCustTblBtn_Click(Nothing, Nothing)
    End Sub

    Private Sub InsertRowBtn_Click(sender As Object, e As EventArgs) Handles InsertRowBtn.Click
        Dim newCId As String = Range("L5").Value
        Dim newFName As String = Range("L6").Value
        Dim newLName As String = Range("L7").Value
        Dim newCity As String = Range("L8").Value
        Dim newState As String = Range("L9").Value

        Dim myString As String = String.Format(
            "INSERT INTO Customer2 (C_id, F_name, L_name, City, State) values ('{0}', '{1}', '{2}', '{3}', '{4}')",
                                    newCId, newFName, newLName, newCity, newState)
        ExecuteNonQuery(myString)

        LoadCustTblBtn_Click(Nothing, Nothing)
    End Sub
End Class
