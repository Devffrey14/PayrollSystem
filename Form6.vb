Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing.Printing

Public Class Form6
    ' Declare connection at class level
    Private connectionString As String = "Data Source=DESKTOP-JA12RFJ\SQLEXPRESS;Initial Catalog=Payroll;Integrated Security=True;Encrypt=False"
    Private printDocument As PrintDocument
    Private printDataTable As DataTable
    Private currentRowIndex As Integer = 0
    Private currentPage As Integer = 1 ' For page numbering

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Initialize form components
            InitializeDataGridView()
            InitializePrintDocument()

            ' Load data
            LoadDataToDataGridView()
        Catch ex As Exception
            MessageBox.Show("Error loading form: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub InitializeDataGridView()
        ' Configure DataGridView properties
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.MultiSelect = False
    End Sub

    Private Sub InitializePrintDocument()
        printDocument = New PrintDocument()
        AddHandler printDocument.PrintPage, AddressOf PrintDocument_PrintPage
        printDocument.DefaultPageSettings.Landscape = True ' Print in landscape mode for better fit
    End Sub

    Private Sub LoadDataToDataGridView(Optional searchTerm As String = "")
        Dim query As String = "SELECT EmployeeID, FirstName, Surname, Position, BasicSalary, " &
                             "OvertimeHours, OvertimePay, NHIL, GETFundLevy, PAYETax, SSNIT, " &
                             "NetSalary, DatePaid FROM Payslips"

        ' Add search condition if search term is provided
        If Not String.IsNullOrWhiteSpace(searchTerm) Then
            query += " WHERE EmployeeID LIKE @SearchTerm OR " &
                     "FirstName LIKE @SearchTerm OR " &
                     "Surname LIKE @SearchTerm"
        End If

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add search parameter if needed
                    If Not String.IsNullOrWhiteSpace(searchTerm) Then
                        command.Parameters.AddWithValue("@SearchTerm", "%" & searchTerm & "%")
                    End If

                    Dim adapter As New SqlDataAdapter(command)
                    Dim table As New DataTable()

                    connection.Open()
                    adapter.Fill(table)

                    ' Bind the DataTable to the DataGridView
                    DataGridView1.DataSource = table
                    printDataTable = table ' Store for printing

                    ' Format columns for better display
                    FormatDataGridViewColumns()
                End Using
            End Using
        Catch ex As SqlException
            MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FormatDataGridViewColumns()
        Try
            ' Format currency columns
            Dim currencyColumns As String() = {"BasicSalary", "OvertimePay", "NHIL", "GETFundLevy", "PAYETax", "SSNIT", "NetSalary"}

            For Each colName As String In currencyColumns
                If DataGridView1.Columns.Contains(colName) Then
                    DataGridView1.Columns(colName).DefaultCellStyle.Format = "C2"
                    DataGridView1.Columns(colName).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
            Next

            ' Format date column
            If DataGridView1.Columns.Contains("DatePaid") Then
                DataGridView1.Columns("DatePaid").DefaultCellStyle.Format = "d"
            End If

            ' Auto-size columns to fit data
            DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        Catch ex As Exception
            MessageBox.Show("Error formatting columns: " & ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
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

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        ' Simple export to CSV (can be opened in Excel)
        Try
            ' Check if there's data to export
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("No data to export!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Create save file dialog
            Using saveFileDialog As New SaveFileDialog()
                saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
                saveFileDialog.Title = "Export Payroll Data"
                saveFileDialog.FileName = "PayrollData_" & DateTime.Now.ToString("yyyyMMdd") & ".csv"

                If saveFileDialog.ShowDialog() = DialogResult.OK Then
                    ' Create a StreamWriter to write to the file
                    Using writer As New StreamWriter(saveFileDialog.FileName)
                        ' Write column headers
                        Dim headers As New List(Of String)
                        For Each column As DataGridViewColumn In DataGridView1.Columns
                            headers.Add("""" & column.HeaderText.Replace("""", """""") & """")
                        Next
                        writer.WriteLine(String.Join(",", headers))

                        ' Write data rows
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            If Not row.IsNewRow Then
                                Dim dataValues As New List(Of String)
                                For Each cell As DataGridViewCell In row.Cells
                                    Dim cellValue As String = If(cell.Value IsNot Nothing, cell.Value.ToString(), "")
                                    dataValues.Add("""" & cellValue.Replace("""", """""") & """")
                                Next
                                writer.WriteLine(String.Join(",", dataValues))
                            End If
                        Next
                    End Using

                    MessageBox.Show("Data exported successfully to: " & saveFileDialog.FileName,
                                    "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error exporting data: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
        Form4.Show()
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("No data to print!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Reset counters before printing
            currentPage = 1
            currentRowIndex = 0

            ' Show print dialog
            Using printDialog As New PrintDialog()
                printDialog.Document = printDocument
                printDialog.PrinterSettings.DefaultPageSettings.Landscape = True

                If printDialog.ShowDialog() = DialogResult.OK Then
                    printDocument.Print()
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error printing document: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PrintDocument_PrintPage(sender As Object, e As PrintPageEventArgs)
        ' Constants for layout
        Dim leftMargin As Integer = e.MarginBounds.Left
        Dim topMargin As Integer = e.MarginBounds.Top
        Dim rightMargin As Integer = e.MarginBounds.Right
        Dim rowHeight As Integer = 30
        Dim currentY As Integer = topMargin
        Dim columnSpacing As Integer = 20 ' Space between columns

        ' Fonts
        Dim headerFont As New Font("Times New Roman", 12, FontStyle.Bold)
        Dim columnFont As New Font("Times New Roman", 9, FontStyle.Bold)
        Dim rowFont As New Font("Times New Roman", 8)
        Dim footerFont As New Font("Times New Roman", 10, FontStyle.Italic)

        ' Calculate scaling factor to fit all columns
        Dim totalRequiredWidth As Integer = 0
        Dim columnWidths As New Dictionary(Of String, Integer)

        ' First pass: Calculate required widths
        For Each column As DataColumn In printDataTable.Columns
            ' Measure column header
            Dim headerWidth As Integer = CInt(e.Graphics.MeasureString(column.ColumnName, columnFont).Width) + columnSpacing

            ' Measure data content
            Dim maxDataWidth As Integer = 0
            For Each row As DataRow In printDataTable.Rows
                If row(column) IsNot DBNull.Value Then
                    Dim cellValue As String = FormatCellValue(row(column), column.ColumnName)
                    Dim dataWidth As Integer = CInt(e.Graphics.MeasureString(cellValue, rowFont).Width) + columnSpacing
                    If dataWidth > maxDataWidth Then maxDataWidth = dataWidth
                End If
            Next

            ' Use the wider of header or data
            columnWidths.Add(column.ColumnName, Math.Max(headerWidth, maxDataWidth))
            totalRequiredWidth += columnWidths(column.ColumnName)
        Next

        ' Calculate scaling factor if needed
        Dim scaleFactor As Single = 1.0F
        Dim availableWidth As Integer = e.MarginBounds.Width

        If totalRequiredWidth > availableWidth Then
            scaleFactor = CSng(availableWidth / totalRequiredWidth) * 0.95 ' 5% margin
        End If

        ' Print header
        e.Graphics.DrawString("Payroll Payslips Report", headerFont, Brushes.Black, leftMargin, currentY)
        currentY += 30

        ' Print date and page info
        e.Graphics.DrawString($"Printed on: {DateTime.Now.ToString("dd/MM/yyyy hh:mm tt")}", rowFont, Brushes.Black, leftMargin, currentY)
        e.Graphics.DrawString($"Page {currentPage}", rowFont, Brushes.Black, rightMargin - 50, currentY)
        currentY += 25

        ' Print column headers
        Dim xPosition As Integer = leftMargin
        For Each column As DataColumn In printDataTable.Columns
            Dim scaledWidth As Integer = CInt(columnWidths(column.ColumnName) * scaleFactor)
            e.Graphics.DrawString(column.ColumnName, columnFont, Brushes.Black, New RectangleF(xPosition, currentY, scaledWidth, rowHeight))
            xPosition += scaledWidth
        Next

        currentY += rowHeight
        e.Graphics.DrawLine(Pens.Black, leftMargin, currentY, xPosition, currentY)
        currentY += 10

        ' Print rows
        Dim rowsPrinted As Integer = 0
        While currentRowIndex < printDataTable.Rows.Count AndAlso currentY + rowHeight < e.MarginBounds.Bottom
            Dim row As DataRow = printDataTable.Rows(currentRowIndex)
            xPosition = leftMargin

            For Each column As DataColumn In printDataTable.Columns
                Dim scaledWidth As Integer = CInt(columnWidths(column.ColumnName) * scaleFactor)
                Dim cellValue As String = If(row(column) IsNot DBNull.Value, FormatCellValue(row(column), column.ColumnName), "")

                e.Graphics.DrawString(cellValue, rowFont, Brushes.Black,
                                     New RectangleF(xPosition, currentY, scaledWidth, rowHeight))
                xPosition += scaledWidth
            Next

            currentY += rowHeight
            currentRowIndex += 1
            rowsPrinted += 1
        End While

        ' Print footer if last page
        If currentRowIndex >= printDataTable.Rows.Count Then
            currentY += 20
            e.Graphics.DrawString($"Total Records: {printDataTable.Rows.Count}", footerFont, Brushes.Black, leftMargin, currentY)
        End If

        ' Set up for next page if needed
        e.HasMorePages = currentRowIndex < printDataTable.Rows.Count
        If e.HasMorePages Then
            currentPage += 1
        Else
            currentPage = 1 ' Reset for next print job
        End If
    End Sub

    Private Function FormatCellValue(value As Object, columnName As String) As String
        ' Format different data types appropriately
        If value Is DBNull.Value Then Return ""

        Select Case columnName
            Case "BasicSalary", "OvertimePay", "NHIL", "GETFundLevy", "PAYETax", "SSNIT", "NetSalary"
                Return String.Format("{0:C2}", Convert.ToDecimal(value))
            Case "DatePaid"
                Return String.Format("{0:d}", Convert.ToDateTime(value))
            Case Else
                Return value.ToString()
        End Select
    End Function

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a Payslip to delete", "Selection Required",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If MessageBox.Show("Are you sure you want to delete this Payslip?", "Confirm Delete",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim employeeID As Integer = CInt(selectedRow.Cells("EmployeeID").Value)

            Dim query As String = "DELETE FROM Payslips WHERE EmployeeID = @EmployeeID"

            Try
                Using connection As New SqlConnection(connectionString)
                    Using command As New SqlCommand(query, connection)
                        command.Parameters.AddWithValue("@EmployeeID", employeeID)
                        connection.Open()
                        Dim rowsAffected As Integer = command.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Payslip deleted successfully!", "Success",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Refresh the DataGridView
                            LoadDataToDataGridView()
                            ' Clear form

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
End Class