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

    Public Function GetADate(query As String) As Date
        Try
            myCommand.CommandText = query
            Return DateTime.Parse(myCommand.ExecuteScalar())
        Catch e As Exception
            MessageBox.Show("Opps! GetADate failed: " + e.Message + " 1/1/1 was returned.")
            Return DateTime.Parse("1/1/1")
        End Try
    End Function

End Module