
Public Class Sheet1

    Private Sub Sheet1_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub FinCalcBtn_Click(sender As Object, e As EventArgs) Handles FinCalcBtn.Click
        Dim UserInput As String = “Please Enter Something”
        Dim InterestRate As Double = 0
        Dim Principal As Double = 0
        Dim NumberOfYears As Double = 0
        Dim Interest As Double = 0
        Dim Sum As Double = 0

        Do
            'Clear designated cells
            Range("A1:C100").Clear()

            'Store interest rate value input by user
            Do
                UserInput = InputBox("Hello! Please enter an interest rate (e.g., 4.725 means 4.725 percent)",
                                 "User input window", "0")
                InterestRate = Double.Parse(UserInput)
            Loop While (InterestRate <= 0) Or (InterestRate > 10)

            InterestRate = InterestRate / 100

            'Store principal value input by user
            Do
                UserInput = InputBox("Please enter a principal (e.g., 1000 - no $)",
                                 "User input window", "0")
                Principal = Double.Parse(UserInput)
            Loop While (Principal <= 0)

            'Store number of years input by user
            Do
                UserInput = InputBox("Please enter the time in years (e.g., 7)", "User input window", "0")
                NumberOfYears = Double.Parse(UserInput)
            Loop While (NumberOfYears < 1) Or (NumberOfYears > 30)

            'Let user choose a case
            UserInput = InputBox("Would you like to see: (a) the interest, (b) the sum of principal + interest, or (c) both a and b (default).",
                             "User input window", "0")

            'Define different cases
            Select Case UserInput
                Case "a"
                    'Prepare the headers for the table
                    Range("A1").Value = "Year"
                    Columns("A").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("A1").Font.Color = Drawing.Color.Red

                    Range("B1").Value = "Interest"
                    Columns("B").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("B1").Font.Color = Drawing.Color.Red

                    'Output the desired values to designated cells
                    Range("A1").Select()
                    For i As Integer = 1 To NumberOfYears Step 1
                        Application.ActiveCell.Offset(i, 0).Value = i
                        Application.ActiveCell.Offset(i, 1).Value = Principal * ((1 + InterestRate) ^ i - 1)
                        Application.ActiveCell.Offset(i, 1).NumberFormat = "$##,##0.00"
                    Next

                Case "b"
                    'Prepare the headers for the table
                    Range("A1").Value = "Year"
                    Columns("A").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("A1").Font.Color = Drawing.Color.Red

                    Range("B1").Value = "Sum of int & prcpl"
                    Columns("B").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("B1").Font.Color = Drawing.Color.Red

                    'Output the desired values to designated cells
                    Range("A1").Select()
                    For i As Integer = 1 To NumberOfYears Step 1
                        Application.ActiveCell.Offset(i, 0).Value = i
                        Application.ActiveCell.Offset(i, 1).Value = Principal * ((1 + InterestRate) ^ i)
                        Application.ActiveCell.Offset(i, 1).NumberFormat = "$##,##0.00"
                    Next

                Case Else
                    'Prepare the headers for the table
                    Range("A1").Value = "Year"
                    Columns("A").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("A1").Font.Color = Drawing.Color.Red

                    Range("B1").Value = "Interest"
                    Columns("B").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("B1").Font.Color = Drawing.Color.Red

                    Range("C1").Value = "Sum of int & prcpl"
                    Columns("C").Select()
                    Application.Selection.Columns.AutoFit()
                    Range("C1").Font.Color = Drawing.Color.Red

                    'Output the desired values to designated cells
                    Range("A1").Select()
                    For i As Integer = 1 To NumberOfYears Step 1
                        Application.ActiveCell.Offset(i, 0).Value = i
                        Application.ActiveCell.Offset(i, 1).Value = Principal * ((1 + InterestRate) ^ i - 1)
                        Application.ActiveCell.Offset(i, 2).Value = Principal * ((1 + InterestRate) ^ i)
                        Application.ActiveCell.Offset(i, 1).NumberFormat = "$##,##0.00"
                        Application.ActiveCell.Offset(i, 2).NumberFormat = "$##,##0.00"
                    Next
            End Select

            'Ask user if he/she wants to calculate again
            UserInput = InputBox("Want to compute some more? (y/n)",
                             "User input window", "n")
        Loop While UserInput = "y"

    End Sub
End Class
