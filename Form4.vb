Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form4
    ' Declare connection at class level
    Private connectionString As String = "Data Source=DESKTOP-JA12RFJ\SQLEXPRESS;Initial Catalog=Payroll;Integrated Security=True;Encrypt=False"

    Dim query As String = "SELECT EmployeeID FROM Employees"
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadEmployeeData()
        ResetForm()

        ' Disable calculated fields
        txtPAYE.Enabled = False
        txtSSNIT.Enabled = False
        txtNHIL.Enabled = False
        txtGFLvy.Enabled = False
        txtNetSalary.Enabled = False
        txtOTPay.Enabled = False

        ' Set default date to today
        DTPpaid.Value = Date.Today
    End Sub

    Private Sub LoadEmployeeData()
        Dim query As String = "SELECT EmployeeID, FirstName, Surname, Position FROM Employees ORDER BY Surname, FirstName"

        Try
            cboEmployeeID.Items.Clear()

            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    connection.Open()
                    Dim reader As SqlDataReader = command.ExecuteReader()

                    ' Create a list to store employee data
                    Dim employees As New List(Of EmployeeInfo)

                    While reader.Read()
                        employees.Add(New EmployeeInfo With {
                            .EmployeeID = Convert.ToInt32(reader("EmployeeID")),
                            .FirstName = reader("FirstName").ToString(),
                            .Surname = reader("Surname").ToString(),
                            .Position = reader("Position").ToString()
                        })
                    End While

                    ' Bind to ComboBox
                    cboEmployeeID.DataSource = employees
                    cboEmployeeID.ValueMember = "EmployeeID"

                    reader.Close()
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show("Database error loading employees: " & ex.Message,
                          "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error loading employee data: " & ex.Message,
                          "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Employee information class
    Private Class EmployeeInfo
        Public Property EmployeeID As Integer
        Public Property FirstName As String
        Public Property Surname As String
        Public Property Position As String

        Public ReadOnly Property FullName As String
            Get
                Return $"{Surname}, {FirstName} (ID: {EmployeeID})"
            End Get
        End Property
    End Class

    Private Sub ResetForm()
        ' Clear combo box selection
        cboEmployeeID.SelectedIndex = -1
        cboEmployeeID.Text = ""

        ' Clear text boxes
        txtFirstName.Clear()
        txtSurname.Clear()
        txtPosition.Clear()
        txtBasicSalary.Clear()
        txtOTHours.Clear()
        txtOTPay.Clear()
        txtNetSalary.Clear()

        ' Reset tax fields to zero
        txtPAYE.Text = "0.00"
        txtSSNIT.Text = "0.00"
        txtNHIL.Text = "0.00"
        txtGFLvy.Text = "0.00"

        ' Reset date to today
        DTPpaid.Value = Date.Today
    End Sub

    Private Sub cboEmployeeID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEmployeeID.SelectedIndexChanged
        If cboEmployeeID.SelectedIndex = -1 Then
            txtFirstName.Clear()
            txtSurname.Clear()
            txtPosition.Clear()
            Exit Sub
        End If

        ' Get the selected employee
        Dim selectedEmployee = DirectCast(cboEmployeeID.SelectedItem, EmployeeInfo)

        ' Populate the text boxes
        txtFirstName.Text = selectedEmployee.FirstName
        txtSurname.Text = selectedEmployee.Surname
        txtPosition.Text = selectedEmployee.Position
    End Sub

    ' Overtime calculation methods
    Private Sub txtOTHours_TextChanged(sender As Object, e As EventArgs) Handles txtOTHours.TextChanged
        If txtOTHours.Focused Then
            CalculateOvertimePay()
        End If
    End Sub

    Private Sub txtBasicSalary_TextChanged(sender As Object, e As EventArgs) Handles txtBasicSalary.TextChanged
        If txtBasicSalary.Focused AndAlso Not String.IsNullOrEmpty(txtOTHours.Text) Then
            CalculateOvertimePay()
        End If
    End Sub

    Private Sub CalculateOvertimePay()
        ' Validate inputs
        If String.IsNullOrWhiteSpace(txtOTHours.Text) OrElse
           String.IsNullOrWhiteSpace(txtBasicSalary.Text) Then
            Exit Sub
        End If

        Try
            ' Parse inputs
            Dim hoursWorked As Decimal = Decimal.Parse(txtOTHours.Text)
            Dim basicSalary As Decimal = Decimal.Parse(txtBasicSalary.Text)

            ' Validate positive values
            If hoursWorked < 0 OrElse basicSalary < 0 Then
                MessageBox.Show("Please enter positive values for hours and salary",
                               "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Calculate overtime pay (time-and-a-half)
            Dim hourlyRate As Decimal = basicSalary / 160D ' 160 hours/month
            Dim overtimePay As Decimal = hoursWorked * hourlyRate * 1.5D

            ' Display result
            txtOTPay.Text = overtimePay.ToString("N2")

        Catch ex As FormatException
            If Not String.IsNullOrEmpty(txtOTHours.Text) AndAlso Not String.IsNullOrEmpty(txtBasicSalary.Text) Then
                MessageBox.Show("Please enter valid numbers for hours and salary",
                               "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Error calculating overtime: " & ex.Message,
                           "Calculation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Net salary calculation
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        ' Validate required fields
        If String.IsNullOrWhiteSpace(txtBasicSalary.Text) Then
            MessageBox.Show("Please enter basic salary", "Input Required",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBasicSalary.Focus()
            Return
        End If

        Try
            ' Get input values
            Dim basicSalary As Decimal = Decimal.Parse(txtBasicSalary.Text)
            Dim overtimePay As Decimal = If(String.IsNullOrWhiteSpace(txtOTPay.Text), 0D, Decimal.Parse(txtOTPay.Text))
            Dim grossSalary As Decimal = basicSalary + overtimePay

            ' Calculate deductions (Ghana-specific rates)
            Dim payeDeduction As Decimal = CalculatePAYE(grossSalary)
            Dim ssnitDeduction As Decimal = grossSalary * 0.055D  ' 5.5%
            Dim nhilDeduction As Decimal = grossSalary * 0.025D   ' 2.5%
            Dim gflvyDeduction As Decimal = grossSalary * 0.025D  ' 2.5%

            ' Calculate net salary
            Dim totalDeductions As Decimal = payeDeduction + ssnitDeduction + nhilDeduction + gflvyDeduction
            Dim netSalary As Decimal = grossSalary - totalDeductions

            ' Display results
            txtPAYE.Text = payeDeduction.ToString("N2")
            txtSSNIT.Text = ssnitDeduction.ToString("N2")
            txtNHIL.Text = nhilDeduction.ToString("N2")
            txtGFLvy.Text = gflvyDeduction.ToString("N2")
            txtNetSalary.Text = netSalary.ToString("N2")

        Catch ex As FormatException
            MessageBox.Show("Please enter valid numeric values", "Input Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error calculating salary: " & ex.Message,
                          "Calculation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function CalculatePAYE(grossSalary As Decimal) As Decimal
        ' Ghana PAYE tax brackets (2023 rates)
        If grossSalary <= 365 Then
            Return 0
        ElseIf grossSalary <= 475 Then
            Return (grossSalary - 365) * 0.05D
        ElseIf grossSalary <= 605 Then
            Return (grossSalary - 475) * 0.1D + 5.5D
        Else
            Return (grossSalary - 605) * 0.175D + 18.5D
        End If
    End Function

    ' Input validation
    Private Sub NumericField_KeyPress(sender As Object, e As KeyPressEventArgs) Handles _
        txtOTHours.KeyPress, txtBasicSalary.KeyPress

        ' Allow only numbers, decimal point, and backspace
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If


    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ResetForm()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Validate required fields
        If cboEmployeeID.SelectedIndex = -1 Then
            MessageBox.Show("Please select an employee", "Input Required",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cboEmployeeID.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtBasicSalary.Text) OrElse
           String.IsNullOrWhiteSpace(txtNetSalary.Text) Then
            MessageBox.Show("Please calculate salary before saving", "Input Required",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm save
        If MessageBox.Show("Save this payslip?", "Confirm Save",
                         MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        ' Prepare SQL command
        Dim query As String = "INSERT INTO Payslips (EmployeeID, FirstName, Surname, Position, " &
                             "BasicSalary, OvertimeHours, OvertimePay, NHIL, GETFundLevy, PAYETax, " &
                             "SSNIT, NetSalary, DatePaid) VALUES (@EmployeeID, @FirstName, @Surname, " &
                             "@Position, @BasicSalary, @OvertimeHours, @OvertimePay, @NHIL, @GETFundLevy, " &
                             "@PAYETax, @SSNIT, @NetSalary, @DatePaid)"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Get selected employee
                    Dim employee = DirectCast(cboEmployeeID.SelectedItem, EmployeeInfo)

                    ' Add parameters
                    command.Parameters.AddWithValue("@EmployeeID", employee.EmployeeID)
                    command.Parameters.AddWithValue("@FirstName", employee.FirstName)
                    command.Parameters.AddWithValue("@Surname", employee.Surname)
                    command.Parameters.AddWithValue("@Position", employee.Position)
                    command.Parameters.AddWithValue("@BasicSalary", Decimal.Parse(txtBasicSalary.Text))
                    command.Parameters.AddWithValue("@OvertimeHours", If(String.IsNullOrEmpty(txtOTHours.Text), 0, Decimal.Parse(txtOTHours.Text)))
                    command.Parameters.AddWithValue("@OvertimePay", If(String.IsNullOrEmpty(txtOTPay.Text), 0, Decimal.Parse(txtOTPay.Text)))
                    command.Parameters.AddWithValue("@NHIL", Decimal.Parse(txtNHIL.Text))
                    command.Parameters.AddWithValue("@GETFundLevy", Decimal.Parse(txtGFLvy.Text))
                    command.Parameters.AddWithValue("@PAYETax", Decimal.Parse(txtPAYE.Text))
                    command.Parameters.AddWithValue("@SSNIT", Decimal.Parse(txtSSNIT.Text))
                    command.Parameters.AddWithValue("@NetSalary", Decimal.Parse(txtNetSalary.Text))
                    command.Parameters.AddWithValue("@DatePaid", DTPpaid.Value.Date)

                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Payslip saved successfully!", "Success",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ResetForm()
                    End If
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show("Database error saving payslip: " & ex.Message,
                           "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error saving payslip: " & ex.Message,
                           "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub btnViewSlips_Click(sender As Object, e As EventArgs) Handles btnViewSlips.Click
        Form6.Show()
        Me.Hide()
    End Sub
End Class