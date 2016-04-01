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
        Dim newCId As String = Range("K5").Value
        Dim newFName As String = Range("K6").Value
        Dim newLName As String = Range("K7").Value
        Dim newCity As String = Range("K8").Value
        Dim newState As String = Range("K9").Value

        Dim myString As String = String.Format(
            "INSERT INTO Customer2 (C_id, F_name, L_name, City, State) values ('{0}', '{1}', '{2}', '{3}', '{4}')",
                                    newCId, newFName, newLName, newCity, newState)
        ExecuteNonQuery(myString)

        LoadCustTblBtn_Click(Nothing, Nothing)
    End Sub
End Class