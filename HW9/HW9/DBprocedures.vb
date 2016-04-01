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
