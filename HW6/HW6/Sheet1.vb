
Public Class Sheet1
    Public Function CleanIt(s As String) As String
        Dim cleaned As String = ""
        Dim temp As Char
        For i As Integer = 0 To (s.Length - 1)
            temp = s.Substring(i, 1)
            If IsNumeric(temp) Then
                cleaned = cleaned + temp
                ' else skip
            End If
        Next
        Return cleaned
    End Function
    Public Function FormatPhNo(s As String) As String
        Return "(" + s.Substring(0, 3) + ")-" + s.Substring(3, 3) + "-" + s.Substring(6, 4)
    End Function
    Private Sub Sheet1_Startup() Handles Me.Startup
        StartBtn.Visible = True
        CleanPhNoBtn.Visible = False
        FormatDatesBtn.Visible = False
        ComputeCagrBtn.Visible = False
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub StartBtn_Click(sender As Object, e As EventArgs) Handles StartBtn.Click

        ' Hide Start Button only
        StartBtn.Visible = False
        CleanPhNoBtn.Visible = True
        FormatDatesBtn.Visible = True
        ComputeCagrBtn.Visible = True

        ' store the address of the current active sheet, i.e., the ‘target’
        Dim myActiveS As Excel.Worksheet = Application.ActiveSheet
        ' select a file
        Dim myFile As String = Application.GetOpenFilename()
        ' get the data in a new temporary workbook
        Application.Workbooks.OpenText(myFile, , , Excel.XlTextParsingType.xlDelimited, , , , , True)
        ' store the address of the temporary workbook
        Dim myActiveWB As Excel.Workbook = Application.ActiveWorkbook
        ' copy the content from the temporary to the ‘target’ sheet
        myActiveS.Range("A1:J1000").Value = Application.ActiveSheet.Range("A1:J1000").Value
        ' close the temp workbook
        myActiveWB.Close()
        'Autofit the column width
        Application.ActiveSheet.Range("A1:J1000").Select()
        Application.Selection.Columns.AutoFit()

    End Sub

    Private Sub CleanPhNoBtn_Click(sender As Object, e As EventArgs) Handles CleanPhNoBtn.Click
        Dim LastRowOfData As Integer
        Dim CleanedPhNo As String = ""
        LastRowOfData = Cells(Rows.Count, "C").End(Excel.XlDirection.xlUp).Row()
        For i As Integer = 1 To LastRowOfData
            CleanedPhNo = CleanIt(Cells(i, "C").value)
            If CleanedPhNo.Length = 10 Then
                Cells(i, "C").Value = FormatPhNo(CleanedPhNo)
            Else
                Cells(i, "C").Interior.Color = System.Drawing.Color.Red
            End If
        Next
    End Sub

    Private Sub FormatDatesBtn_Click(sender As Object, e As EventArgs) Handles FormatDatesBtn.Click
        Dim LastRow As Integer = 0
        For Each col As String In {"D", "G", "I"}
            LastRow = Cells(Rows.Count, col).End(Excel.XlDirection.xlUp).Row
            For i As Integer = 1 To LastRow
                Cells(i, col).NumberFormat = "mm/dd/yyyy"
                If IsDate(Cells(i, col).Value) Then
                    Cells(i, col).Interior.Color = System.Drawing.Color.White
                Else
                    Cells(i, col).Interior.Color = System.Drawing.Color.Yellow
                End If
            Next
        Next
    End Sub

    Private Sub ComputeCagrBtn_Click(sender As Object, e As EventArgs) Handles ComputeCagrBtn.Click

    End Sub
End Class
