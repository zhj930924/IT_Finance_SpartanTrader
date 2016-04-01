
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
            Range("A1:B2").Clear()

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
                    Interest = Principal * ((1 + InterestRate) ^ NumberOfYears - 1)
                    Range("A1").Value = Interest.ToString("C2")
                    Range("A1").ColumnWidth = 15
                    Range("B1").Value = "is the interest that you requested"
                Case "b"
                    Interest = Principal * ((1 + InterestRate) ^ NumberOfYears - 1)
                    Sum = Principal + Interest
                    Range("A1").Value = Sum.ToString("C2")
                    Range("A1").ColumnWidth = 15
                    Range("B1").Value = "is the sum of interest and principal"
                Case Else
                    Interest = Principal * ((1 + InterestRate) ^ NumberOfYears - 1)
                    Sum = Principal + Interest
                    'Show interest
                    Range("A1").Value = Interest.ToString("C2")
                    Range("A1").ColumnWidth = 15
                    Range("B1").Value = "is the interest that you requested"
                    'Show sum
                    Range("A2").Value = Sum.ToString("C2")
                    Range("A2").ColumnWidth = 15
                    Range("B2").Value = "is the sum of interest and principal"
            End Select

            'Ask user if he/she wants to calculate again
            UserInput = InputBox("Want to compute some more? (y/n)",
                             "User input window", "n")
        Loop While UserInput = "y"

    End Sub
End Class
