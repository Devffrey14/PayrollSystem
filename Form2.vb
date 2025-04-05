Public Class Form2
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        Form7.Show()
        Me.Hide()
    End Sub

    Private Sub btnEmployees_Click(sender As Object, e As EventArgs) Handles btnEmployees.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Form4.Show()
        Me.Hide()
    End Sub
End Class