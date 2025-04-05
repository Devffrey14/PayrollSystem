Imports System.Data.SqlClient

Public Class Form3
    ' Declare connection at class level so it can be used throughout the form
    Private connectionString As String = "Data Source=DESKTOP-JA12RFJ\SQLEXPRESS;Initial Catalog=Payroll;Integrated Security=True;Encrypt=False"
    Private connection As SqlConnection
    Private currentEmployeeID As Integer = -1 ' To track which employee is being edited

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize form controls
        InitializeFormControls()

        ' Load data into DataGridView
        LoadDataToDataGridView()
    End Sub

    Private Sub InitializeFormControls()
        txtAge.Enabled = False

        cboTitle.Items.AddRange({"Mr.", "Mrs.", "Ms.", "Dr."})
        cboGender.Items.AddRange({"Male", "Female"})
        cboPosition.Items.AddRange({
            "Driver", "Tech Support", "Data Analyst", "Safety Officer",
            "First Aid", "Secretary", "Caterer", "Accountant",
            "Manager", "Quality Assurance"
        })
    End Sub

    Private Sub LoadDataToDataGridView(Optional searchTerm As String = "")
        Dim query As String = "SELECT EmployeeID, Title, FirstName, Surname, Gender, DOB, Age, Position, DateHired, Telephone, Email, Image FROM Employees"

        ' Add search filter if search term is provided
        If Not String.IsNullOrWhiteSpace(searchTerm) Then
            query += " WHERE FirstName LIKE @SearchTerm OR Surname LIKE @SearchTerm OR Position LIKE @SearchTerm"
        End If

        Try
            connection = New SqlConnection(connectionString)
            Dim adapter As New SqlDataAdapter(query, connection)

            ' Add search parameter if needed
            If Not String.IsNullOrWhiteSpace(searchTerm) Then
                adapter.SelectCommand.Parameters.AddWithValue("@SearchTerm", "%" & searchTerm & "%")
            End If

            Dim table As New DataTable()

            connection.Open()
            adapter.Fill(table)

            ' Bind the DataTable to the DataGridView
            DataGridView1.DataSource = table

            ' Optional: Auto-size columns to fit data
            DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        Catch ex As SqlException
            MessageBox.Show("SQL Error: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("General Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If connection IsNot Nothing AndAlso connection.State <> ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub dtpDOB_ValueChanged(sender As Object, e As EventArgs) Handles dtpDOB.ValueChanged
        Dim Today As Integer = Date.Today.Year
        Dim DOB As Integer = dtpDOB.Value.Year
        txtAge.Text = (Today - DOB).ToString()
    End Sub

    Private Sub btnloadPhoto_Click(sender As Object, e As EventArgs) Handles btnloadPhoto.Click
        Dim OPF As New OpenFileDialog With {
            .Filter = "Image Files (*.jpg; *.jpeg; *.png; *.JPG)|*.jpg;*.jpeg;*.png;*.JPG"
        }

        If OPF.ShowDialog() = DialogResult.OK Then
            Try
                PictureBox1.Image = Image.FromFile(OPF.FileName)
                MessageBox.Show("Photo Loaded Successfully", "Information Desk", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error loading image: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' Validate required fields
        If String.IsNullOrWhiteSpace(txtFirstName.Text) OrElse
           String.IsNullOrWhiteSpace(txtSurname.Text) OrElse
           cboGender.SelectedIndex = -1 OrElse
           cboPosition.SelectedIndex = -1 Then
            MessageBox.Show("Please fill in all required fields", "Validation Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Validate age is numeric
        Dim age As Integer
        If Not Integer.TryParse(txtAge.Text, age) Then
            MessageBox.Show("Please enter a valid age", "Validation Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Create SQL command with parameters
        Dim query As String = "INSERT INTO Employees (Title, FirstName, Surname, Gender, DOB, Age, Position, DateHired, Telephone, Email, Image) " &
                             "VALUES (@Title, @FirstName, @Surname, @Gender, @DOB, @Age, @Position, @DateHired, @Telephone, @Email, @Image)"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@Title", If(cboTitle.SelectedItem IsNot Nothing, cboTitle.SelectedItem.ToString(), DBNull.Value))
                    command.Parameters.AddWithValue("@FirstName", txtFirstName.Text.Trim())
                    command.Parameters.AddWithValue("@Surname", txtSurname.Text.Trim())
                    command.Parameters.AddWithValue("@Gender", cboGender.SelectedItem.ToString())
                    command.Parameters.AddWithValue("@DOB", dtpDOB.Value.Date)
                    command.Parameters.AddWithValue("@Age", age)
                    command.Parameters.AddWithValue("@Position", cboPosition.SelectedItem.ToString())
                    command.Parameters.AddWithValue("@DateHired", DTPHired.Value.Date)
                    command.Parameters.AddWithValue("@Telephone", If(String.IsNullOrWhiteSpace(txtTelephone.Text), DBNull.Value, txtTelephone.Text.Trim()))
                    command.Parameters.AddWithValue("@Email", If(String.IsNullOrWhiteSpace(txtEmail.Text), DBNull.Value, txtEmail.Text.Trim()))

                    ' Add image parameter
                    If PictureBox1.Image IsNot Nothing Then
                        Dim ms As New System.IO.MemoryStream()
                        PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
                        command.Parameters.Add("@Image", SqlDbType.VarBinary).Value = ms.ToArray()
                    Else
                        command.Parameters.AddWithValue("@Image", DBNull.Value)
                    End If

                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Employee added successfully!", "Success",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Refresh the DataGridView
                        LoadDataToDataGridView()
                        ' Clear form for next entry
                        ClearForm()
                    Else
                        MessageBox.Show("No records were added", "Information",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show($"Database error: {ex.Message}{vbCrLf}Please check your data and try again.",
                           "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearForm()
        ' Clear all input fields
        cboTitle.SelectedIndex = -1
        txtFirstName.Clear()
        txtSurname.Clear()
        cboGender.SelectedIndex = -1
        dtpDOB.Value = DateTime.Now
        txtAge.Clear()
        cboPosition.SelectedIndex = -1
        DTPHired.Value = DateTime.Now
        txtTelephone.Clear()
        txtEmail.Clear()
        currentEmployeeID = -1 ' Reset the tracking ID

        ' Clear the image
        If PictureBox1.Image IsNot Nothing Then
            PictureBox1.Image.Dispose()
            PictureBox1.Image = Nothing
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Store the current employee ID for editing
            currentEmployeeID = Convert.ToInt32(selectedRow.Cells("EmployeeID").Value)

            ' Populate form fields with selected row data
            cboTitle.SelectedItem = If(selectedRow.Cells("Title").Value Is DBNull.Value, Nothing, selectedRow.Cells("Title").Value.ToString())
            txtFirstName.Text = If(selectedRow.Cells("FirstName").Value Is DBNull.Value, String.Empty, selectedRow.Cells("FirstName").Value.ToString())
            txtSurname.Text = If(selectedRow.Cells("Surname").Value Is DBNull.Value, String.Empty, selectedRow.Cells("Surname").Value.ToString())
            cboGender.SelectedItem = If(selectedRow.Cells("Gender").Value Is DBNull.Value, Nothing, selectedRow.Cells("Gender").Value.ToString())
            dtpDOB.Value = If(selectedRow.Cells("DOB").Value Is DBNull.Value, DateTime.Now, Convert.ToDateTime(selectedRow.Cells("DOB").Value))
            txtAge.Text = If(selectedRow.Cells("Age").Value Is DBNull.Value, String.Empty, selectedRow.Cells("Age").Value.ToString())
            cboPosition.SelectedItem = If(selectedRow.Cells("Position").Value Is DBNull.Value, Nothing, selectedRow.Cells("Position").Value.ToString())
            DTPHired.Value = If(selectedRow.Cells("DateHired").Value Is DBNull.Value, DateTime.Now, Convert.ToDateTime(selectedRow.Cells("DateHired").Value))
            txtTelephone.Text = If(selectedRow.Cells("Telephone").Value Is DBNull.Value, String.Empty, selectedRow.Cells("Telephone").Value.ToString())
            txtEmail.Text = If(selectedRow.Cells("Email").Value Is DBNull.Value, String.Empty, selectedRow.Cells("Email").Value.ToString())

            ' Handle the image
            Try
                ' Clear current image if exists
                If PictureBox1.Image IsNot Nothing Then
                    PictureBox1.Image.Dispose()
                    PictureBox1.Image = Nothing
                End If

                ' Load new image if available
                If Not selectedRow.Cells("Image").Value Is DBNull.Value Then
                    Dim imageData As Byte() = DirectCast(selectedRow.Cells("Image").Value, Byte())
                    Using ms As New System.IO.MemoryStream(imageData)
                        PictureBox1.Image = Image.FromStream(ms)
                    End Using
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                Else
                    PictureBox1.Image = Nothing ' Or set a default image
                End If
            Catch ex As Exception
                MessageBox.Show("Error loading image: " & ex.Message, "Image Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
                PictureBox1.Image = Nothing
            End Try
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to delete", "Selection Required",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If MessageBox.Show("Are you sure you want to delete this record?", "Confirm Delete",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim employeeID As Integer = CInt(selectedRow.Cells("EmployeeID").Value)

            Dim query As String = "DELETE FROM Employees WHERE EmployeeID = @EmployeeID"

            Try
                Using connection As New SqlConnection(connectionString)
                    Using command As New SqlCommand(query, connection)
                        command.Parameters.AddWithValue("@EmployeeID", employeeID)
                        connection.Open()
                        Dim rowsAffected As Integer = command.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Employee deleted successfully!", "Success",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Refresh the DataGridView
                            LoadDataToDataGridView()
                            ' Clear form
                            ClearForm()
                        Else
                            MessageBox.Show("No records were deleted", "Information",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                End Using
            Catch ex As SqlException
                MessageBox.Show($"Database error: {ex.Message}{vbCrLf}Cannot delete this record.",
                               "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
                MessageBox.Show($"Error: {ex.Message}", "Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim searchTerm As String = txtSearch.Text.Trim()
            LoadDataToDataGridView(searchTerm)
        Catch ex As Exception
            MessageBox.Show("Search error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearch_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSearch.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            btnSearch.PerformClick()
            e.Handled = True
        End If
    End Sub

    Private Sub btnClearSearch_Click(sender As Object, e As EventArgs) Handles btnClearSearch.Click
        Try
            txtSearch.Clear()
            LoadDataToDataGridView()
        Catch ex As Exception
            MessageBox.Show("Error clearing search: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        ' Check if an employee is selected
        If currentEmployeeID = -1 Then
            MessageBox.Show("Please select an employee to edit", "Selection Required",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Validate required fields
        If String.IsNullOrWhiteSpace(txtFirstName.Text) OrElse
           String.IsNullOrWhiteSpace(txtSurname.Text) OrElse
           cboGender.SelectedIndex = -1 OrElse
           cboPosition.SelectedIndex = -1 Then
            MessageBox.Show("Please fill in all required fields", "Validation Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Validate age is numeric
        Dim age As Integer
        If Not Integer.TryParse(txtAge.Text, age) Then
            MessageBox.Show("Please enter a valid age", "Validation Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Create SQL command with parameters
        Dim query As String = "UPDATE Employees SET " &
                             "Title = @Title, " &
                             "FirstName = @FirstName, " &
                             "Surname = @Surname, " &
                             "Gender = @Gender, " &
                             "DOB = @DOB, " &
                             "Age = @Age, " &
                             "Position = @Position, " &
                             "DateHired = @DateHired, " &
                             "Telephone = @Telephone, " &
                             "Email = @Email, " &
                             "Image = @Image " &
                             "WHERE EmployeeID = @EmployeeID"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@EmployeeID", currentEmployeeID)
                    command.Parameters.AddWithValue("@Title", If(cboTitle.SelectedItem IsNot Nothing, cboTitle.SelectedItem.ToString(), DBNull.Value))
                    command.Parameters.AddWithValue("@FirstName", txtFirstName.Text.Trim())
                    command.Parameters.AddWithValue("@Surname", txtSurname.Text.Trim())
                    command.Parameters.AddWithValue("@Gender", cboGender.SelectedItem.ToString())
                    command.Parameters.AddWithValue("@DOB", dtpDOB.Value.Date)
                    command.Parameters.AddWithValue("@Age", age)
                    command.Parameters.AddWithValue("@Position", cboPosition.SelectedItem.ToString())
                    command.Parameters.AddWithValue("@DateHired", DTPHired.Value.Date)
                    command.Parameters.AddWithValue("@Telephone", If(String.IsNullOrWhiteSpace(txtTelephone.Text), DBNull.Value, txtTelephone.Text.Trim()))
                    command.Parameters.AddWithValue("@Email", If(String.IsNullOrWhiteSpace(txtEmail.Text), DBNull.Value, txtEmail.Text.Trim()))

                    ' Add image parameter
                    If PictureBox1.Image IsNot Nothing Then
                        Dim ms As New System.IO.MemoryStream()
                        PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
                        command.Parameters.Add("@Image", SqlDbType.VarBinary).Value = ms.ToArray()
                    Else
                        command.Parameters.AddWithValue("@Image", DBNull.Value)
                    End If

                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Employee updated successfully!", "Success",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Refresh the DataGridView
                        LoadDataToDataGridView()
                        ' Clear form for next entry
                        ClearForm()
                    Else
                        MessageBox.Show("No records were updated", "Information",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show($"Database error: {ex.Message}{vbCrLf}Please check your data and try again.",
                           "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class